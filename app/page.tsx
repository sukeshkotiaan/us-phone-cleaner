'use client';
import { useState, useRef, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { processPhone, normalizeForSupp, PhoneResult } from '@/lib/phoneCleaner';

type RowResult = { row: string[]; results: Record<number, PhoneResult> };

const STATUS_COLORS: Record<string, string> = {
  clean:      'bg-green-100 text-green-800',
  stripped:   'bg-yellow-100 text-yellow-800',
  invalid:    'bg-red-100 text-red-800',
  junk:       'bg-red-100 text-red-800',
  blank:      'bg-gray-100 text-gray-500',
  suppressed: 'bg-purple-100 text-purple-800',
};
const STATUS_LABELS: Record<string, string> = {
  clean:'Clean', stripped:'Stripped 1', invalid:'Invalid',
  junk:'Junk/Fake', blank:'Blank', suppressed:'Suppressed',
};

function Stat({ label, value, color, sub }: { label: string; value: number; color: string; sub?: string }) {
  return (
    <div className="bg-white rounded-xl border border-gray-200 p-4">
      <div className={`text-2xl font-semibold ${color}`}>{value.toLocaleString()}</div>
      <div className="text-xs text-gray-500 mt-1">{label}</div>
      {sub && <div className="text-xs text-gray-400 mt-0.5">{sub}</div>}
    </div>
  );
}

function Badge({ status }: { status: string }) {
  return (
    <span className={`inline-flex items-center px-2 py-0.5 rounded-full text-xs font-medium ${STATUS_COLORS[status] || 'bg-gray-100 text-gray-600'}`}>
      {STATUS_LABELS[status] || status}
    </span>
  );
}

export default function Home() {
  const [headers, setHeaders] = useState<string[]>([]);
  const [rawData, setRawData] = useState<string[][]>([]);
  const [processedRows, setProcessedRows] = useState<RowResult[]>([]);
  const [selectedCols, setSelectedCols] = useState<Set<number>>(new Set());
  const [suppNumbers, setSuppNumbers] = useState<string[]>([]);
  const [leadFile, setLeadFile] = useState('');
  const [suppFile, setSuppFile] = useState('');
  const [isDraggingLead, setIsDraggingLead] = useState(false);
  const [isDraggingSupp, setIsDraggingSupp] = useState(false);
  const [processing, setProcessing] = useState(false);
  const [showSupp, setShowSupp] = useState(false);
  const [numverifyKey, setNumverifyKey] = useState('');
  const [nvProgress, setNvProgress] = useState(0);
  const [nvRunning, setNvRunning] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [showRules, setShowRules] = useState(false);
  const leadInputRef = useRef<HTMLInputElement>(null);
  const suppInputRef = useRef<HTMLInputElement>(null);

  const readWorkbook = (file: File): Promise<{ headers: string[]; rows: string[][] }> =>
    new Promise(res => {
      const reader = new FileReader();
      reader.onload = e => {
        const wb = XLSX.read(e.target!.result, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }) as string[][];
        res({ headers: json[0]?.map(String) || [], rows: json.slice(1) });
      };
      reader.readAsArrayBuffer(file);
    });

  const runProcessing = useCallback((rows: string[][], cols: Set<number>, supp: string[]) => {
    setProcessing(true);
    const suppSet = new Set(supp);
    const result: RowResult[] = rows.map(row => {
      const results: Record<number, PhoneResult> = {};
      cols.forEach(ci => { results[ci] = processPhone(String(row[ci] ?? ''), suppSet); });
      return { row, results };
    });
    setProcessedRows(result);
    setProcessing(false);
  }, []);

  const handleLeadFile = async (file: File) => {
    const { headers: h, rows } = await readWorkbook(file);
    setHeaders(h); setRawData(rows); setLeadFile(file.name);
    setProcessedRows([]); setSelectedCols(new Set());
    const rx = /phone|mobile|cell|tel|contact|number|ph\b|mob/i;
    const auto = new Set<number>();
    h.forEach((col, i) => { if (rx.test(col)) auto.add(i); });
    setSelectedCols(auto);
    if (auto.size > 0) runProcessing(rows, auto, suppNumbers);
  };

  const handleSuppFile = async (file: File) => {
    const { headers: h, rows } = await readWorkbook(file);
    const allRows = h.length ? [h, ...rows] : rows;
    const nums: string[] = [];
    allRows.forEach(r => { if (r[0]) nums.push(normalizeForSupp(String(r[0]))); });
    setSuppNumbers(nums);
    setSuppFile(`${file.name} — ${nums.length.toLocaleString()} numbers`);
    if (selectedCols.size > 0 && rawData.length > 0) runProcessing(rawData, selectedCols, nums);
  };

  const toggleCol = (i: number) => {
    const next = new Set(selectedCols);
    next.has(i) ? next.delete(i) : next.add(i);
    setSelectedCols(next);
    if (next.size > 0 && rawData.length > 0) runProcessing(rawData, next, suppNumbers);
  };

  const runNumverify = async () => {
    if (!numverifyKey || !processedRows.length) return;
    setNvRunning(true); setNvProgress(0);
    const updated = [...processedRows];
    const colArr = [...selectedCols];
    const toCheck: { ri: number; ci: number; phone: string }[] = [];
    updated.forEach((r, ri) => {
      colArr.forEach(ci => {
        const res = r.results[ci];
        if (res?.status === 'clean' || res?.status === 'stripped')
          toCheck.push({ ri, ci, phone: res.cleaned });
      });
    });
    for (let i = 0; i < toCheck.length; i++) {
      const { ri, ci, phone } = toCheck[i];
      try {
        const resp = await fetch('/api/numverify', {
          method: 'POST', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ phone, apiKey: numverifyKey }),
        });
        const data = await resp.json();
        updated[ri].results[ci] = { ...updated[ri].results[ci], lineType: data.line_type || '', carrier: data.carrier || '' };
      } catch { /* skip */ }
      setNvProgress(Math.round((i + 1) / toCheck.length * 100));
      await new Promise(r => setTimeout(r, 260));
    }
    setProcessedRows(updated); setNvRunning(false);
  };

  const downloadFile = (mode: 'all' | 'clean') => {
    const colArr = [...selectedCols].sort((a, b) => a - b);
    const extraH = colArr.flatMap(ci => [
      `${headers[ci]||`Col${ci+1}`}_cleaned`,
      `${headers[ci]||`Col${ci+1}`}_status`,
      `${headers[ci]||`Col${ci+1}`}_state_code`,
      `${headers[ci]||`Col${ci+1}`}_state`,
      `${headers[ci]||`Col${ci+1}`}_line_type`,
      `${headers[ci]||`Col${ci+1}`}_carrier`,
    ]);
    let rows = processedRows;
    if (mode === 'clean')
      rows = processedRows.filter(r => colArr.every(ci => r.results[ci]?.status === 'clean' || r.results[ci]?.status === 'stripped'));
    const data = rows.map(({ row, results }) => [
      ...row,
      ...colArr.flatMap(ci => {
        const r = results[ci];
        return [r?.cleaned||'', r?.status||'', r?.stateCode||'', r?.stateName||'', r?.lineType||'', r?.carrier||''];
      })
    ]);
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([[...headers, ...extraH], ...data]);
    XLSX.utils.book_append_sheet(wb, ws, 'Cleaned');
    XLSX.writeFile(wb, mode === 'clean' ? 'leads_dialable.xlsx' : 'leads_all_status.xlsx');
  };

  const colArr = [...selectedCols];
  const stats = { total:0, clean:0, stripped:0, invalid:0, junk:0, blank:0, suppressed:0 };
  processedRows.forEach(({ results }) => {
    stats.total++;
    let worst = 'clean';
    colArr.forEach(ci => {
      const s = results[ci]?.status || 'blank';
      if (['invalid','junk','blank','suppressed'].includes(s)) worst = s;
      else if (s === 'stripped' && worst === 'clean') worst = 'stripped';
    });
    stats[worst as keyof typeof stats]++;
  });

  const hasResults = processedRows.length > 0;
  const dialable = stats.clean + stats.stripped;
  const hasLineType = hasResults && colArr.some(ci => processedRows[0]?.results[ci]?.lineType);

  return (
    <div className="min-h-screen bg-gray-50">
      {/* Header */}
      <div className="bg-white border-b border-gray-200 px-6 py-4 flex items-center justify-between">
        <div className="flex items-center gap-3">
          <div className="w-9 h-9 bg-blue-600 rounded-xl flex items-center justify-center">
            <svg className="w-5 h-5 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M3 5a2 2 0 012-2h3.28a1 1 0 01.948.684l1.498 4.493a1 1 0 01-.502 1.21l-2.257 1.13a11.042 11.042 0 005.516 5.516l1.13-2.257a1 1 0 011.21-.502l4.493 1.498a1 1 0 01.684.949V19a2 2 0 01-2 2h-1C9.716 21 3 14.284 3 7V5z"/>
            </svg>
          </div>
          <div>
            <h1 className="text-base font-semibold text-gray-900">US Phone Cleaner</h1>
            <p className="text-xs text-gray-400">Clean · Validate · State · Line Type</p>
          </div>
        </div>
        <div className="flex items-center gap-2">
          <button onClick={() => setShowRules(!showRules)} className="text-sm text-gray-500 hover:text-gray-700 px-3 py-1.5 rounded-lg hover:bg-gray-100 transition">
            {showRules ? 'Hide rules' : 'View rules'}
          </button>
          <button onClick={() => setShowSettings(!showSettings)} className="text-sm text-gray-500 hover:text-gray-700 px-3 py-1.5 rounded-lg hover:bg-gray-100 transition flex items-center gap-1.5">
            <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.065 2.572c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 00-2.572 1.065c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.065-2.572c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37.996.608 2.296.07 2.572-1.065z"/>
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 12a3 3 0 11-6 0 3 3 0 016 0z"/>
            </svg>
            API Settings
          </button>
        </div>
      </div>

      <div className="max-w-6xl mx-auto px-6 py-6 space-y-5">

        {/* Rules */}
        {showRules && (
          <div className="bg-white rounded-xl border border-gray-200 p-5">
            <h3 className="text-sm font-semibold text-gray-700 mb-3">Cleaning rules applied</h3>
            <div className="grid grid-cols-2 gap-x-8 text-sm">
              <div>
                <p className="text-xs font-semibold text-green-700 uppercase tracking-wide mb-2">Keep</p>
                {['10-digit numbers → kept as-is','11-digit starting with 1 → strip the 1, output 10 digits'].map(r=>(
                  <div key={r} className="flex gap-2 text-gray-600 mb-1"><span className="text-green-500">✓</span>{r}</div>
                ))}
              </div>
              <div>
                <p className="text-xs font-semibold text-red-700 uppercase tracking-wide mb-2">Remove</p>
                {['Fewer than 10 digits','More than 11 digits','11-digit not starting with 1','All same digit (0000000000…9999999999)','Sequential (1234567890, 9876543210)','Area code starts with 0 or 1','Exchange (digits 4–6) starts with 0 or 1','555-01xx Hollywood fakes','Repeating 2-digit blocks (1212121212)','Repeating 3-digit blocks (1231231234)','Known placeholders'].map(r=>(
                  <div key={r} className="flex gap-2 text-gray-600 mb-1"><span className="text-red-500">✗</span>{r}</div>
                ))}
              </div>
            </div>
          </div>
        )}

        {/* API Settings */}
        {showSettings && (
          <div className="bg-white rounded-xl border border-gray-200 p-5">
            <h3 className="text-sm font-semibold text-gray-700 mb-1">Numverify — Line Type Detection</h3>
            <p className="text-xs text-gray-400 mb-4">Identifies mobile, landline, or VoIP. Free tier: 100 calls/month at <a href="https://numverify.com" target="_blank" rel="noreferrer" className="text-blue-500 hover:underline">numverify.com</a></p>
            <div className="flex items-end gap-4">
              <div className="flex-1">
                <label className="text-xs text-gray-500 block mb-1">API Key</label>
                <input type="password" placeholder="Paste your Numverify API key" value={numverifyKey}
                  onChange={e => setNumverifyKey(e.target.value)}
                  className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"/>
              </div>
            </div>
            {numverifyKey && hasResults && (
              <div className="mt-4 pt-4 border-t border-gray-100">
                {nvRunning ? (
                  <div>
                    <div className="flex justify-between text-xs text-gray-500 mb-1.5"><span>Looking up line types…</span><span>{nvProgress}%</span></div>
                    <div className="w-full bg-gray-100 rounded-full h-1.5"><div className="bg-blue-500 h-1.5 rounded-full transition-all" style={{width:`${nvProgress}%`}}/></div>
                  </div>
                ) : (
                  <button onClick={runNumverify} className="bg-blue-600 text-white text-sm px-4 py-2 rounded-lg hover:bg-blue-700 transition">
                    Run lookup on {dialable.toLocaleString()} dialable numbers
                  </button>
                )}
              </div>
            )}
          </div>
        )}

        {/* Upload panels */}
        <div className="grid grid-cols-2 gap-4">
          {/* Lead file — primary */}
          <div
            className={`border-2 border-dashed rounded-2xl p-10 text-center cursor-pointer transition-all ${
              isDraggingLead ? 'border-blue-400 bg-blue-50' :
              leadFile ? 'border-green-300 bg-green-50' :
              'border-gray-200 bg-white hover:border-blue-300 hover:bg-blue-50'
            }`}
            onClick={() => leadInputRef.current?.click()}
            onDragOver={e => { e.preventDefault(); setIsDraggingLead(true); }}
            onDragLeave={() => setIsDraggingLead(false)}
            onDrop={e => { e.preventDefault(); setIsDraggingLead(false); const f = e.dataTransfer.files[0]; if (f) handleLeadFile(f); }}
          >
            <input ref={leadInputRef} type="file" accept=".csv,.xlsx,.xls" className="hidden"
              onChange={e => { if (e.target.files?.[0]) handleLeadFile(e.target.files[0]); }}/>
            <div className="text-4xl mb-3">{leadFile ? '✅' : '📂'}</div>
            {leadFile ? (
              <>
                <p className="text-sm font-semibold text-green-700">{leadFile}</p>
                <p className="text-xs text-green-600 mt-1">{rawData.length.toLocaleString()} rows · {headers.length} columns</p>
                <p className="text-xs text-gray-400 mt-2">Click to replace</p>
              </>
            ) : (
              <>
                <p className="text-sm font-medium text-gray-600">Drop your lead file here</p>
                <p className="text-xs text-gray-400 mt-1">CSV or XLSX · any column layout</p>
              </>
            )}
          </div>

          {/* Suppression — optional, collapsible */}
          <div className="flex flex-col gap-3">
            <button
              onClick={() => setShowSupp(!showSupp)}
              className={`w-full flex items-center justify-between px-4 py-3 rounded-xl border transition text-sm font-medium ${
                suppFile ? 'border-purple-200 bg-purple-50 text-purple-700' : 'border-gray-200 bg-white text-gray-500 hover:bg-gray-50'
              }`}
            >
              <div className="flex items-center gap-2">
                <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m5.618-4.016A11.955 11.955 0 0112 2.944a11.955 11.955 0 01-8.618 3.04A12.02 12.02 0 003 9c0 5.591 3.824 10.29 9 11.622 5.176-1.332 9-6.03 9-11.622 0-1.042-.133-2.052-.382-3.016z"/>
                </svg>
                {suppFile ? suppFile : 'Add suppression list (optional)'}
              </div>
              <svg className={`w-4 h-4 transition-transform ${showSupp ? 'rotate-180' : ''}`} fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7"/>
              </svg>
            </button>

            {showSupp && (
              <div
                className={`border-2 border-dashed rounded-xl p-6 text-center cursor-pointer transition-all flex-1 ${
                  isDraggingSupp ? 'border-purple-400 bg-purple-50' :
                  suppFile ? 'border-purple-200 bg-purple-50' :
                  'border-gray-200 bg-white hover:border-purple-300 hover:bg-purple-50'
                }`}
                onClick={() => suppInputRef.current?.click()}
                onDragOver={e => { e.preventDefault(); setIsDraggingSupp(true); }}
                onDragLeave={() => setIsDraggingSupp(false)}
                onDrop={e => { e.preventDefault(); setIsDraggingSupp(false); const f = e.dataTransfer.files[0]; if (f) handleSuppFile(f); }}
              >
                <input ref={suppInputRef} type="file" accept=".csv,.xlsx,.xls" className="hidden"
                  onChange={e => { if (e.target.files?.[0]) handleSuppFile(e.target.files[0]); }}/>
                <div className="text-3xl mb-2">{suppFile ? '🛡️' : '📋'}</div>
                {suppFile ? (
                  <p className="text-xs text-purple-600">Click to replace</p>
                ) : (
                  <>
                    <p className="text-sm text-gray-500">Drop suppression file</p>
                    <p className="text-xs text-gray-400 mt-1">Single-column list · CSV or XLSX</p>
                  </>
                )}
              </div>
            )}
          </div>
        </div>

        {/* Processing */}
        {processing && (
          <div className="bg-white rounded-xl border border-blue-100 p-4 flex items-center gap-3">
            <div className="w-4 h-4 border-2 border-blue-500 border-t-transparent rounded-full animate-spin"/>
            <span className="text-sm text-blue-700">Cleaning and validating numbers…</span>
          </div>
        )}

        {/* Column selector */}
        {headers.length > 0 && !processing && (
          <div className="bg-white rounded-xl border border-gray-200 p-4">
            <div className="text-xs font-medium text-gray-400 uppercase tracking-wide mb-3">Select phone column(s) to clean</div>
            <div className="flex flex-wrap gap-2">
              {headers.map((h, i) => (
                <button key={i} onClick={() => toggleCol(i)}
                  className={`px-3 py-1.5 rounded-full text-sm border transition ${selectedCols.has(i) ? 'bg-blue-100 border-blue-300 text-blue-800 font-medium' : 'bg-gray-50 border-gray-200 text-gray-500 hover:border-gray-300'}`}>
                  {h || `Col ${i + 1}`}
                </button>
              ))}
            </div>
          </div>
        )}

        {/* Stats */}
        {hasResults && !processing && (
          <div className={`grid gap-3 ${suppNumbers.length > 0 ? 'grid-cols-6' : 'grid-cols-5'}`}>
            <Stat label="Total rows" value={stats.total} color="text-gray-800"/>
            <Stat label="Clean 10-digit" value={stats.clean} color="text-green-600"/>
            <Stat label="Stripped leading 1" value={stats.stripped} color="text-yellow-600"/>
            <Stat label="Invalid / blank" value={stats.invalid + stats.blank} color="text-red-500"/>
            <Stat label="Junk / fake" value={stats.junk} color="text-red-500"/>
            {suppNumbers.length > 0 && <Stat label="Suppressed" value={stats.suppressed} color="text-purple-600"/>}
          </div>
        )}

        {/* Preview + download */}
        {hasResults && !processing && (
          <div className="bg-white rounded-xl border border-gray-200 overflow-hidden">
            <div className="flex items-center justify-between px-4 py-3 border-b border-gray-100">
              <span className="text-sm font-medium text-gray-700">
                Preview <span className="text-gray-400 font-normal">(first 15 rows)</span>
              </span>
              <div className="flex gap-2">
                <button onClick={() => downloadFile('all')}
                  className="text-sm border border-gray-200 px-3 py-1.5 rounded-lg hover:bg-gray-50 transition text-gray-600">
                  ↓ All rows with status
                </button>
                <button onClick={() => downloadFile('clean')}
                  className="text-sm bg-blue-600 text-white px-4 py-1.5 rounded-lg hover:bg-blue-700 transition font-medium">
                  ↓ Dialable only ({dialable.toLocaleString()})
                </button>
              </div>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead className="bg-gray-50 text-xs text-gray-500 uppercase tracking-wide">
                  <tr>
                    {headers.slice(0, 4).map((h, i) => (
                      <th key={i} className={`px-4 py-2.5 text-left font-medium whitespace-nowrap ${selectedCols.has(i) ? 'bg-blue-50 text-blue-600' : ''}`}>{h || `Col ${i + 1}`}</th>
                    ))}
                    {colArr.sort((a,b)=>a-b).map(ci => (
                      <th key={`c${ci}`} className="px-4 py-2.5 text-left font-medium bg-green-50 text-green-700 whitespace-nowrap">
                        {headers[ci]||`Col ${ci+1}`} cleaned
                      </th>
                    ))}
                    <th className="px-4 py-2.5 text-left font-medium bg-green-50 text-green-700 whitespace-nowrap">State</th>
                    {hasLineType && <th className="px-4 py-2.5 text-left font-medium bg-purple-50 text-purple-700 whitespace-nowrap">Line type</th>}
                    {hasLineType && <th className="px-4 py-2.5 text-left font-medium bg-purple-50 text-purple-700 whitespace-nowrap">Carrier</th>}
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {processedRows.slice(0, 15).map((r, ri) => (
                    <tr key={ri} className="hover:bg-gray-50">
                      {headers.slice(0, 4).map((_, ci) => (
                        <td key={ci} className={`px-4 py-2.5 text-gray-700 max-w-xs truncate ${selectedCols.has(ci) ? 'bg-blue-50/30' : ''}`}>
                          <span className={selectedCols.has(ci) ? 'font-mono text-xs text-gray-400' : ''}>{String(r.row[ci] || '')}</span>
                        </td>
                      ))}
                      {colArr.map(ci => {
                        const res = r.results[ci];
                        return (
                          <td key={`r${ci}`} className="px-4 py-2.5 bg-green-50/30">
                            <div className="flex items-center gap-2">
                              <span className="font-mono text-xs font-medium text-gray-800">{res?.cleaned || '—'}</span>
                              <Badge status={res?.status || 'blank'}/>
                            </div>
                            {res?.reason && !['clean','stripped'].includes(res.status) && (
                              <div className="text-xs text-gray-400 mt-0.5 truncate max-w-48">{res.reason}</div>
                            )}
                          </td>
                        );
                      })}
                      <td className="px-4 py-2.5 bg-green-50/30 text-xs text-gray-600 whitespace-nowrap">
                        {colArr[0] !== undefined ? (r.results[colArr[0]]?.stateName || <span className="text-gray-300">—</span>) : '—'}
                      </td>
                      {hasLineType && (
                        <td className="px-4 py-2.5 bg-purple-50/30">
                          <span className={`text-xs font-medium capitalize ${
                            colArr.map(ci => r.results[ci]?.lineType).filter(Boolean)[0] === 'mobile' ? 'text-green-700' :
                            colArr.map(ci => r.results[ci]?.lineType).filter(Boolean)[0] === 'landline' ? 'text-gray-600' : 'text-orange-600'
                          }`}>
                            {colArr.map(ci => r.results[ci]?.lineType).filter(Boolean)[0] || '—'}
                          </span>
                        </td>
                      )}
                      {hasLineType && (
                        <td className="px-4 py-2.5 bg-purple-50/30 text-xs text-gray-600">
                          {colArr.map(ci => r.results[ci]?.carrier).filter(Boolean)[0] || '—'}
                        </td>
                      )}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
