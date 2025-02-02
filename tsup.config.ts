import { defineConfig } from 'tsup';

export default defineConfig({
  entry: {
    exceljs: 'src/index.js',
  },
  clean: true,
  format: ['esm', 'cjs'],
});
