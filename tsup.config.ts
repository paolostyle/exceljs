import { defineConfig } from 'tsup';

export default defineConfig({
  entry: ['lib/index.js'],
  clean: true,
  format: ['esm', 'cjs'],
});
