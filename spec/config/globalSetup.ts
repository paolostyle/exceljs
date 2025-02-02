import { mkdir, rm } from 'node:fs/promises';

const outputDir = './spec/out';

export async function setup() {
  await mkdir(outputDir, { recursive: true });
}

export async function teardown() {
  await rm(outputDir, { recursive: true, force: true });
}
