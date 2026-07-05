import { defineConfig } from 'vite';

// Base is env-driven (AUDIT P8): Vercel (primary) serves at '/', the GitHub
// Pages mirror sets VITE_BASE=/btc-ifc-viewer/ in its workflow.
export default defineConfig({
    base: process.env.VITE_BASE ?? '/',
    root: './src',
    publicDir: '../public',
    build: {
        outDir: '../dist',
        emptyOutDir: true,
    },
    server: {
        port: 3001,
        open: true,
    },
    optimizeDeps: {
        exclude: ['web-ifc'],
    },
});
