import { NextRequest, NextResponse } from 'next/server';

export async function GET() {
  return NextResponse.json({ loaded: false, size: 0 });
}

export async function POST(req: NextRequest) {
  try {
    const { phones, suppNumbers } = await req.json();
    if (!Array.isArray(phones)) return NextResponse.json({ error: 'phones must be an array' }, { status: 400 });
    const set = new Set<string>(suppNumbers || []);
    const results: Record<string, boolean> = {};
    phones.forEach((p: string) => { results[p] = set.has(p); });
    return NextResponse.json({ results });
  } catch {
    return NextResponse.json({ error: 'Failed' }, { status: 500 });
  }
}
