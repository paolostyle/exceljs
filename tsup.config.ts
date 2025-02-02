import { defineConfig } from 'tsup';

export default defineConfig({
  entry: ['src/index.js'],
  clean: true,
  format: ['esm', 'cjs'],
});
