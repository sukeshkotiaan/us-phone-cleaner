import { getState } from './areaCodeData';

export type PhoneStatus = 'clean' | 'stripped' | 'invalid' | 'junk' | 'blank' | 'suppressed';

export interface PhoneResult {
  original: string;
  cleaned: string;
  status: PhoneStatus;
  reason: string;
  stateCode: string;
  stateName: string;
  lineType: string;
  carrier: string;
}

const KNOWN_JUNK = new Set([
  '1234567890','0123456789','9876543210','1231234567','1231231234',
  '0000000000','1111111111','2222222222','3333333333','4444444444',
  '5555555555','6666666666','7777777777','8888888888','9999999999',
  '1234512345','5432154321','1212121212','0101010101','1010101010',
]);

export function validateAndClean(raw: string): { digits: string; status: PhoneStatus; reason: string } {
  if (!raw && raw !== 0 as unknown) return { digits: '', status: 'blank', reason: 'Empty value' };
  const digits = String(raw).replace(/[^\d]/g, '');
  if (!digits || digits.length < 10) return { digits: '', status: 'invalid', reason: `Too short (${digits.length} digits)` };
  if (digits.length > 11) return { digits: '', status: 'invalid', reason: `Too long (${digits.length} digits)` };

  let ten = digits;
  let wasStripped = false;

  if (digits.length === 11) {
    if (digits[0] !== '1') return { digits: '', status: 'invalid', reason: '11-digit number not starting with 1' };
    ten = digits.slice(1);
    wasStripped = true;
  }

  // Junk checks on 10-digit number
  if (KNOWN_JUNK.has(ten)) return { digits: '', status: 'junk', reason: 'Known placeholder/fake number' };
  if (/^(\d)\1{9}$/.test(ten)) return { digits: '', status: 'junk', reason: 'All same digit' };
  if (ten === '1234567890' || ten === '9876543210' || ten === '0123456789') return { digits: '', status: 'junk', reason: 'Sequential digits' };
  if (ten[0] === '0' || ten[0] === '1') return { digits: '', status: 'junk', reason: `Invalid area code (starts with ${ten[0]})` };
  if (ten[3] === '0' || ten[3] === '1') return { digits: '', status: 'junk', reason: `Invalid exchange (starts with ${ten[3]})` };
  if (ten.slice(0, 3) === '555' && ten.slice(3, 5) === '01') return { digits: '', status: 'junk', reason: '555-01xx fake number' };

  // Repeating 2-digit block: ABABABABAB
  const b2 = ten.slice(0, 2);
  if (b2.repeat(5) === ten) return { digits: '', status: 'junk', reason: 'Repeating 2-digit pattern' };

  // Repeating 3-digit block: ABCABCABCA
  const b3 = ten.slice(0, 3);
  if ((b3 + b3 + b3 + ten[9]) === ten) return { digits: '', status: 'junk', reason: 'Repeating 3-digit pattern' };

  // Repeating 4-digit block
  const b4 = ten.slice(0, 4);
  if (ten.slice(0, 4) === b4 && ten.slice(4, 8) === b4 && ten.slice(8, 10) === b4.slice(0, 2)) {
    return { digits: '', status: 'junk', reason: 'Repeating 4-digit pattern' };
  }

  return { digits: ten, status: wasStripped ? 'stripped' : 'clean', reason: wasStripped ? 'Stripped leading 1' : '' };
}

export function processPhone(raw: string, suppSet: Set<string>, numverifyData?: { line_type?: string; carrier?: string }): PhoneResult {
  const val = validateAndClean(raw);

  if (val.status === 'blank' || val.status === 'invalid' || val.status === 'junk') {
    return {
      original: String(raw),
      cleaned: '',
      status: val.status,
      reason: val.reason,
      stateCode: '',
      stateName: '',
      lineType: '',
      carrier: '',
    };
  }

  const ten = val.digits;

  // Suppression check
  if (suppSet.size > 0 && suppSet.has(ten)) {
    return {
      original: String(raw),
      cleaned: ten,
      status: 'suppressed',
      reason: 'On suppression list',
      stateCode: '',
      stateName: '',
      lineType: '',
      carrier: '',
    };
  }

  const { code, name } = getState(ten);

  return {
    original: String(raw),
    cleaned: ten,
    status: val.status,
    reason: val.reason,
    stateCode: code,
    stateName: name,
    lineType: numverifyData?.line_type || '',
    carrier: numverifyData?.carrier || '',
  };
}

export function normalizeForSupp(val: string): string {
  const s = String(val).replace(/[^\d]/g, '');
  if (s.length === 11 && s[0] === '1') return s.slice(1);
  return s;
}
