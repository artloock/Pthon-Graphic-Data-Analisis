import basicSsl from "@vitejs/plugin-basic-ssl";
import { defineConfig } from "vite";

export default defineConfig({
  plugins: [basicSsl()],
  build: { rollupOptions: { input: "taskpane.html" } },
  server: { port: 3000, strictPort: true },
  preview: { port: 3000, strictPort: true }
});
