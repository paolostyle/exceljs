import { defineConfig } from 'tsup';

export default defineConfig({
  entry: ['lib/exceljs.nodejs.js'],
  clean: true,
  format: ['esm', 'cjs'],
});
