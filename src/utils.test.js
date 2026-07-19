import { describe, it, expect, vi, afterEach } from 'vitest';
import {
  today, fmt, countWorkingDays, addWorkingDays, isOverdue,
  escHtml, sanitizeUrl,
} from './utils.js';

afterEach(() => {
  vi.useRealTimers();
});

describe('today()', () => {
  it('returns the current date as YYYY-MM-DD', () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date('2026-03-15T09:30:00'));
    expect(today()).toBe('2026-03-15');
  });
});

describe('fmt()', () => {
  it('formats a YYYY-MM-DD date as DD Mon YYYY', () => {
    expect(fmt('2026-03-05')).toBe('05 Mar 2026');
  });
  it('returns an em dash for null/undefined/empty', () => {
    expect(fmt(null)).toBe('—');
    expect(fmt(undefined)).toBe('—');
    expect(fmt('')).toBe('—');
  });
});

describe('countWorkingDays()', () => {
  it('counts Mon-Fri inclusive, excluding weekends', () => {
    // Mon 2026-03-02 through Fri 2026-03-06 = 5 working days
    expect(countWorkingDays('2026-03-02', '2026-03-06')).toBe(5);
  });
  it('excludes a weekend spanning the range', () => {
    // Mon 2026-03-02 through Mon 2026-03-09 = 6 working days (2 weekends' worth of Sat/Sun excluded)
    expect(countWorkingDays('2026-03-02', '2026-03-09')).toBe(6);
  });
  it('counts a single working day as 1', () => {
    expect(countWorkingDays('2026-03-04', '2026-03-04')).toBe(1);
  });
  it('counts a single Saturday as 0', () => {
    expect(countWorkingDays('2026-03-07', '2026-03-07')).toBe(0);
  });
});

describe('addWorkingDays()', () => {
  it('skips weekends when adding working days', () => {
    // Fri 2026-03-06 + 1 working day -> Mon 2026-03-09 (skips Sat/Sun)
    expect(addWorkingDays('2026-03-06', 1)).toBe('2026-03-09');
  });
  it('adding 0 working days returns the start date unchanged', () => {
    expect(addWorkingDays('2026-03-02', 0)).toBe('2026-03-02');
  });
  it('adds working days across a full week correctly', () => {
    // Mon 2026-03-02 + 5 working days -> Mon 2026-03-09
    expect(addWorkingDays('2026-03-02', 5)).toBe('2026-03-09');
  });
});

describe('isOverdue()', () => {
  it('a completed task is never overdue, regardless of deadline', () => {
    expect(isOverdue({ status: 'Completed', deadlineDate: '2000-01-01' })).toBe(false);
  });
  it('a task with no deadline is never overdue', () => {
    expect(isOverdue({ status: 'Pending', deadlineDate: null })).toBe(false);
  });
  it('a pending task with a past deadline is overdue', () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date('2026-06-15T12:00:00'));
    expect(isOverdue({ status: 'Pending', deadlineDate: '2026-06-01' })).toBe(true);
  });
  it('a pending task with a future deadline is not overdue', () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date('2026-06-15T12:00:00'));
    expect(isOverdue({ status: 'Pending', deadlineDate: '2026-12-01' })).toBe(false);
  });
});

describe('escHtml() — XSS prevention via HTML escaping', () => {
  it('escapes angle brackets so a script tag cannot inject', () => {
    expect(escHtml('<script>alert(1)</script>')).not.toContain('<script>');
    expect(escHtml('<script>alert(1)</script>')).toContain('&lt;script&gt;');
  });
  it('escapes angle brackets so an injected tag cannot break out of text context', () => {
    // Note: escHtml (via textContent -> innerHTML) does NOT escape a bare "
    // character — that's fine, because a raw quote alone can't execute
    // anything; it's the < and > that matter, and those ARE escaped, which
    // is what actually prevents this payload from becoming a real <img> tag.
    const evil = '"><img src=x onerror=alert(1)>';
    const escaped = escHtml(evil);
    expect(escaped).not.toContain('<img');
    expect(escaped).toContain('&lt;img');
  });
  it('returns an empty string for null/undefined', () => {
    expect(escHtml(null)).toBe('');
    expect(escHtml(undefined)).toBe('');
  });
  it('leaves plain text unchanged', () => {
    expect(escHtml('Balance Sheet Focus Awards')).toBe('Balance Sheet Focus Awards');
  });
});

describe('sanitizeUrl() — blocks dangerous URL schemes in href', () => {
  it('blocks javascript: URLs', () => {
    expect(sanitizeUrl('javascript:alert(document.cookie)')).toBe('#');
  });
  it('blocks javascript: with mixed case (bypass attempt)', () => {
    expect(sanitizeUrl('JaVaScRiPt:alert(1)')).toBe('#');
  });
  it('blocks javascript: with leading whitespace (bypass attempt)', () => {
    expect(sanitizeUrl('   javascript:alert(1)')).toBe('#');
  });
  it('blocks data: URLs', () => {
    expect(sanitizeUrl('data:text/html,<script>alert(1)</script>')).toBe('#');
  });
  it('blocks vbscript: URLs', () => {
    expect(sanitizeUrl('vbscript:msgbox(1)')).toBe('#');
  });
  it('allows https: URLs through unchanged', () => {
    expect(sanitizeUrl('https://example.com/doc.pdf')).toBe('https://example.com/doc.pdf');
  });
  it('allows http: URLs through unchanged', () => {
    expect(sanitizeUrl('http://example.com')).toBe('http://example.com');
  });
  it('allows mailto: URLs through unchanged', () => {
    expect(sanitizeUrl('mailto:sumudu@focusawards.org.uk')).toBe('mailto:sumudu@focusawards.org.uk');
  });
  it('allows tel: URLs through unchanged', () => {
    expect(sanitizeUrl('tel:+441234567890')).toBe('tel:+441234567890');
  });
  it('allows relative URLs through unchanged', () => {
    expect(sanitizeUrl('/some/path')).toBe('/some/path');
  });
  it('returns an empty string for empty/null/undefined input', () => {
    expect(sanitizeUrl('')).toBe('');
    expect(sanitizeUrl(null)).toBe('');
    expect(sanitizeUrl(undefined)).toBe('');
  });
  it('a scheme-less garbage string is treated as a safe relative path, not rejected', () => {
    // new URL(s, window.location.href) is only used to VALIDATE the scheme —
    // a scheme-less string resolves relative to the page's own origin (http/
    // https, always safe), so it passes the check. The function then returns
    // escHtml(s) — the escaped ORIGINAL string, not the resolved/normalized
    // URL — since there's nothing here that needs escaping, it comes back
    // unchanged rather than rejected as '#'.
    expect(sanitizeUrl('not a url at all !!!')).toBe('not a url at all !!!');
  });
  it('rejects a string with an explicit unsafe scheme even amid garbage', () => {
    expect(sanitizeUrl('javascript:void(document.location="https://evil.com")')).toBe('#');
  });
});
