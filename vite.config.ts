import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import path from "path";

const apiServerUrl = process.env.API_SERVER_URL ?? "http://localhost:3000";
let apiTarget: URL;
try {
  apiTarget = new URL(apiServerUrl);
} catch {
  throw new Error("API_SERVER_URL must be an absolute http:// or https:// URL.");
}
if (
  !["http:", "https:"].includes(apiTarget.protocol) ||
  apiTarget.username ||
  apiTarget.password ||
  apiTarget.search ||
  apiTarget.hash
) {
  throw new Error("API_SERVER_URL must use HTTP or HTTPS and must not contain credentials, a query, or a fragment.");
}

export default defineConfig({
  plugins: [react()],
  resolve: {
    alias: { "@": path.resolve(import.meta.dirname, "./src") },
  },
  server: {
    host: true,
    port: 5002,
    strictPort: true,
    proxy: {
      "/api": {
        target: apiTarget.toString().replace(/\/$/, ""),
        changeOrigin: true,
        configure(proxy) {
          proxy.on("error", (_error, _request, response) => {
            if (!("writeHead" in response) || !("end" in response)) return;
            response.writeHead(503, { "content-type": "application/json; charset=utf-8" });
            response.end(JSON.stringify({
              error: "خادم API غير متصل. شغّل خادم API وتأكد من أن API_SERVER_URL يشير إلى عنوانه.",
            }));
          });
        },
      },
    },
  },
  build: {
    rolldownOptions: {
      output: {
        codeSplitting: {
          groups: [
            {
              name: "spreadsheet",
              test: /[\\/]node_modules[\\/]xlsx[\\/]/,
              priority: 20,
            },
            {
              name: "vendor",
              test: /[\\/]node_modules[\\/]/,
              maxSize: 350_000,
              priority: 10,
            },
          ],
        },
      },
    },
  },
});
