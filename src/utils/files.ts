export const loadFileToStream = async (
  path: string,
): Promise<ReadableStream<Uint8Array>> => {
  if (typeof Bun !== 'undefined') {
    return Bun.file(path).stream();
  }
  if (typeof Deno !== 'undefined') {
    return (await Deno.open(path)).readable;
  }
  if (typeof window !== 'undefined') {
    const res = await fetch(path);

    if (!res.body) {
      throw new Error(`Failed to fetch ${path}`);
    }

    return res.body;
  }

  const fs = await import('node:fs');
  const { ReadableStream } = await import('node:stream/web');
  const nodeStream = fs.createReadStream(path);
  return ReadableStream.from(
    nodeStream,
  ) as unknown as ReadableStream<Uint8Array>;
};

export const writeFileFromStream = async (
  stream: ReadableStream<Uint8Array>,
  path: string,
): Promise<void> => {
  if (typeof Bun !== 'undefined') {
    await Bun.write(path, new Response(stream));
    return;
  }
  if (typeof Deno !== 'undefined') {
    const file = await Deno.open(path, { write: true, create: true });
    return await stream.pipeTo(file.writable);
  }
  if (typeof window !== 'undefined') {
    const a = document.createElement('a');
    const blob = await new Response(stream).blob();
    const url = URL.createObjectURL(blob);
    a.href = url;
    a.download = path;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 0);
  }

  const { Readable } = await import('node:stream');
  const fs = await import('node:fs');
  const { mkdir } = await import('node:fs/promises');
  const { dirname } = await import('node:path');

  await mkdir(dirname(path), { recursive: true });
  const fileStream = Readable.from(stream);
  const writeStream = fs.createWriteStream(path);
  fileStream.pipe(writeStream);
};
