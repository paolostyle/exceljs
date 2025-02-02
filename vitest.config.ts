import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    globals: true,
    globalSetup: ['spec/config/globalSetup.ts'],
    setupFiles: ['spec/config/vitestSetup.ts'],
  },
});
