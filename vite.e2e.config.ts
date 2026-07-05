import { defineConfig, mergeConfig } from 'vite';

import baseConfig from './vite.config';

// E2E artifact config (AUDIT T4): identical to the production build except for
// the explicit VITE_E2E define that compiles the window.__viewer/__world test
// hooks in. Playwright builds with this config and serves it via
// `vite preview`, so the suite tests the real minified artifact.
export default mergeConfig(
  baseConfig,
  defineConfig({
    define: {
      'import.meta.env.VITE_E2E': JSON.stringify('true'),
    },
    preview: {
      host: '127.0.0.1',
      port: 4173,
      strictPort: true,
    },
    server: {
      host: '127.0.0.1',
      open: false,
      port: 4173,
      strictPort: true,
    },
  }),
);
