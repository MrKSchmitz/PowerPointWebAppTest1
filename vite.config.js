import { defineConfig } from "vite";
import { getHttpsServerOptions } from "office-addin-dev-certs";

// The manifest points at https://localhost:3000, so the port is fixed.
export default defineConfig(async ({ command }) => ({
    root: "src",
    build: {
        outDir: "../dist",
        emptyOutDir: true,
        rollupOptions: {
            input: "src/taskpane.html"
        }
    },
    server: {
        port: 3000,
        strictPort: true,
        https: command === "serve" ? await getHttpsServerOptions() : undefined
    }
}));
