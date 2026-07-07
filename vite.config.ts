import { resolve } from 'node:path';

import { defineConfig, type Plugin } from 'vite';

import { pwaPlugin } from './scripts/pwa-plugin.mjs';

// Dev-server-only CSP relax (A1): Vite's HMR runs over ws:// which CSP3
// 'self' does not cover. The built artifact keeps the strict policy from
// index.html — this transform never applies to `vite build`.
const devCspRelax = (): Plugin => ({
    name: 'dev-csp-relax',
    apply: 'serve',
    transformIndexHtml(html) {
        return html.replace("connect-src 'self'", "connect-src 'self' ws: wss:");
    },
});

// Base is env-driven (AUDIT P8): Vercel (primary) serves at '/', the GitHub
// Pages mirror sets VITE_BASE=/btc-ifc-viewer/ in its workflow.
export default defineConfig({
    base: process.env.VITE_BASE ?? '/',
    plugins: [devCspRelax(), pwaPlugin()],
    root: './src',
    publicDir: '../public',
    build: {
        outDir: '../dist',
        emptyOutDir: true,
        // MPA (W4.1): the full app (index.html) + the chromeless embed
        // (embed.html) are separate entry points sharing core/* modules.
        rollupOptions: {
            input: {
                main: resolve(__dirname, 'src/index.html'),
                embed: resolve(__dirname, 'src/embed.html'),
            },
            output: {
                // W5.1 (P1): split the heavy engine deps into separate, long-lived
                // cacheable chunks. Each is content-hashed independently, so an app
                // code change no longer invalidates the ~5.7MB engine download.
                // `web-ifc` (the ~5.9MB embind glue) is only reached via the IFC
                // conversion path — the conversion runs in a dedicated worker
                // (workers/ifc-conversion.worker.ts, spawned from
                // core/ifc-conversion-client.ts), which Vite bundles as its own
                // chunk, so web-ifc lands off the initial shell.
                manualChunks(id: string) {
                    if (id.includes('node_modules/web-ifc/')) return 'web-ifc';
                    if (id.includes('node_modules/three/')) return 'three';
                    if (id.includes('node_modules/@thatopen/')) return 'thatopen';
                    return undefined;
                },
            },
        },
    },
    server: {
        port: 3001,
        open: true,
    },
    optimizeDeps: {
        exclude: ['web-ifc'],
    },
});
