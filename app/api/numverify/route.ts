import { NextRequest, NextResponse } from 'next/server';

export async function POST(req: NextRequest) {
  const { phone, apiKey } = await req.json();
  if (!phone || !apiKey) return NextResponse.json({ error: 'Missing phone or apiKey' }, { status: 400 });

  try {
    const url = `http://apilayer.net/api/validate?access_key=${apiKey}&number=1${phone}&country_code=US&format=1`;
    const res = await fetch(url);
    const data = await res.json();
    if (!data.valid) return NextResponse.json({ line_type: 'unknown', carrier: '' });
    return NextResponse.json({ line_type: data.line_type || 'unknown', carrier: data.carrier || '' });
  } catch {
    return NextResponse.json({ error: 'Numverify request failed' }, { status: 500 });
  }
}
