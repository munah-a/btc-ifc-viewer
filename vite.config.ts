import { defineConfig, Plugin } from 'vite';

function wasmMimePlugin(): Plugin {
    return {
        name: 'wasm-mime-type',
        configureServer(server) {
            server.middlewares.use((req, res, next) => {
                if (req.url?.endsWith('.wasm')) {
                    res.setHeader('Content-Type', 'application/wasm');
                }
                next();
            });
        },
    };
}

export default defineConfig({
    base: '/btc-ifc-viewer/',
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
    plugins: [wasmMimePlugin()],
});
