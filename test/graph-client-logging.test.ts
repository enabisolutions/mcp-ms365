import { describe, expect, it } from 'vitest';
import { describeOptionsForLog } from '../src/graph-client.js';

describe('describeOptionsForLog', () => {
  it('never includes the access token, but records that one was present', () => {
    const out = describeOptionsForLog({
      method: 'GET',
      accessToken: 'eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.secret.sig',
    });
    expect(out).not.toContain('eyJ0eXAi');
    expect(out).not.toContain('secret');
    expect(out).toContain('"method":"GET"');
    expect(out).toContain('[accessToken=REDACTED]');
  });

  it('adds no marker when there was no token', () => {
    expect(describeOptionsForLog({ method: 'GET' })).toBe('{"method":"GET"}');
  });
});
