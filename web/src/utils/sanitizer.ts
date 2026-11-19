/**
 * Data sanitization and defensive conversion utilities.
 * Fixes known defects: '$-50.00' sign loss, alphanumeric '1E2' corruption, and multi-PO hyphen splitting.
 */

export function cleanString(val: unknown): string {
  if (val === null || val === undefined) return '';
  const s = String(val).replace(/\u00a0/g, ' ').trim();
  const low = s.toLowerCase();
  if (['nan', 'none', 'null', '#n/a', '#value!', '#ref!', '#div/0!', '#name?'].includes(low)) {
    return '';
  }
  return s.replace(/\s+/g, ' ');
}

export function cleanMaterialCode(val: unknown): string {
  if (val === null || val === undefined) return '';
  let s = String(val).trim().replace(/\u00a0/g, '');
  if (!s || ['nan', 'none', '#n/a', '#value!', '#ref!'].includes(s.toLowerCase())) return '';

  // Strictly decode true numeric scientific notation (e.g. 1.455102001E+09), never alphanumeric like '1E2'
  if (/^\d+\.\d+[eE][+-]?\d+$/.test(s) || /^\d+[eE][+-]\d+$/.test(s)) {
    const f = Number(s);
    if (!Number.isNaN(f) && Number.isFinite(f)) {
      s = Math.round(f).toLocaleString('fullwide', { useGrouping: false });
    }
  }

  // Strip trailing .0 from numeric Excel values
  if (/^\d+\.0+$/.test(s)) {
    s = s.replace(/\.0+$/, '');
  }

  // Keep alphanumeric, underscore, and hyphen
  s = s.replace(/[^\w-]/g, '');

  // Strip SAP 18-char leading zero padding
  if (/^\d+$/.test(s) && s.length > 10 && s.startsWith('000')) {
    s = s.replace(/^0+/, '');
  }

  return s;
}

export function cleanNumber(val: unknown, defaultVal = 0): number {
  if (val === null || val === undefined) return defaultVal;
  if (typeof val === 'number') {
    return Number.isNaN(val) || !Number.isFinite(val) ? defaultVal : val;
  }

  let s = String(val).trim().replace(/\u00a0/g, '');
  if (!s || ['nan', 'none', 'null', '-', 'c', 'x', '#n/a', '#value!', '#ref!', '#div/0!'].includes(s.toLowerCase())) {
    return defaultVal;
  }

  // Detect sign BEFORE stripping currency symbols (fixes '$-50.00' -> -50)
  let isNegative = false;
  if (s.startsWith('(') && s.endsWith(')')) {
    isNegative = true;
    s = s.slice(1, -1).trim();
  } else if (s.includes('-')) {
    if (/^[€$₡£\s]*-/.test(s)) {
      isNegative = true;
    }
  }

  s = s.replace(/[^\d.,]/g, '');
  if (!s) return defaultVal;

  if (s.includes(',') && s.includes('.')) {
    if (s.lastIndexOf(',') > s.lastIndexOf('.')) {
      s = s.replace(/\./g, '').replace(',', '.');
    } else {
      s = s.replace(/,/g, '');
    }
  } else if (s.includes(',')) {
    const parts = s.split(',');
    if (parts.length === 2 && (parts[1].length === 1 || parts[1].length === 2)) {
      s = s.replace(',', '.');
    } else {
      s = s.replace(/,/g, '');
    }
  }

  const num = Number(s);
  if (Number.isNaN(num) || !Number.isFinite(num)) return defaultVal;
  return isNegative ? -num : num;
}

export function cleanInt(val: unknown, defaultVal = 0): number {
  const f = cleanNumber(val, defaultVal);
  if (!Number.isFinite(f)) return defaultVal;
  return Math.round(f);
}

export function parseMultiPo(poStr: unknown): string[] {
  const cleaned = cleanString(poStr);
  if (!cleaned) return [];
  return cleaned
    .split(/[,;/]+/)
    .map((t) => t.trim())
    .filter(Boolean);
}

export function formatExactInt(val: number): string {
  const rounded = Math.round(val);
  return Number.isFinite(rounded) ? String(rounded) : '0';
}
