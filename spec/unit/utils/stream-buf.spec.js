import fs from 'node:fs';
import path from 'node:path';
import { vi } from 'vitest';
import StreamBuf from '#lib/utils/stream-buf.js';
import StringBuf from '#lib/utils/string-buf.js';

describe('StreamBuf', () => {
  // StreamBuf is designed as a general-purpose writable-readable stream
  // However its use in ExcelJS is primarily as a memory buffer between
  // the streaming writers and the archive, hence the tests here will
  // focus just on that.
  it('writes strings as UTF8', () => {
    const stream = new StreamBuf();
    stream.write('Hello, World!');
    const chunk = stream.read();
    expect(chunk instanceof Buffer).toBeTruthy();
    expect(chunk.toString('UTF8')).to.equal('Hello, World!');
  });

  it('writes StringBuf chunks', () => {
    const stream = new StreamBuf();
    const strBuf = new StringBuf({ size: 64 });
    strBuf.addText('Hello, World!');
    stream.write(strBuf);
    const chunk = stream.read();
    expect(chunk instanceof Buffer).toBeTruthy();
    expect(chunk.toString('UTF8')).to.equal('Hello, World!');
  });

  it('signals end', () => {
    const cb = vi.fn();
    const stream = new StreamBuf();
    stream.on('finish', cb);
    stream.write('Hello, World!');
    stream.end();

    expect(cb).toHaveBeenCalledTimes(1);
  });

  it('handles buffers', () =>
    new Promise((resolve, reject) => {
      const s = fs.createReadStream(path.join(__dirname, 'data/image1.png'));
      const sb = new StreamBuf();
      sb.on('finish', () => {
        const buf = sb.toBuffer();
        expect(buf.length).to.equal(1672);
        resolve();
      });
      sb.on('error', reject);
      s.pipe(sb);
    }));
  it('handle unsupported type of chunk', async () => {
    const stream = new StreamBuf();
    try {
      await stream.write({});
      expect.fail('should fail for given argument');
    } catch (e) {
      expect(e.message).to.equal(
        'Chunk must be one of type String, Buffer or StringBuf.',
      );
    }
  });
});
