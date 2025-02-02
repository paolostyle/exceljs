export function bufferToString(chunk) {
  if (typeof chunk === 'string') {
    return chunk;
  }
  return new TextDecoder().decode(chunk);
}
