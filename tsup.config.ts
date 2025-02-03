import { defineConfig } from 'tsup';
import RawPlugin from 'unplugin-raw/esbuild';

export default defineConfig({
  entry: {
    exceljs: 'src/index.js',
  },
  clean: true,
  format: ['esm', 'cjs'],
  plugins: [RawPlugin()],
});
