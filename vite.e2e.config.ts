import { defineConfig, mergeConfig } from 'vite';

import baseConfig from './vite.config';

export default mergeConfig(
  baseConfig,
  defineConfig({
    server: {
      host: '127.0.0.1',
      open: false,
      port: 4173,
      strictPort: true,
    },
  }),
);
