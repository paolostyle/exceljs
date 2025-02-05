import { Buffer } from 'node:buffer';

export function stringToBuffer(str) {
  if (typeof str !== 'string') return str;

  return Buffer.from(new TextEncoder().encode(str).buffer);
}
