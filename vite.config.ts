import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  root: ".",
  plugins: [react()],
  server: {
    middlewareMode: true,
    watch: {
      ignored: ["**/data/settlements/**", "**/data/tmp/**", "**/outputs/**"]
    }
  },
  build: {
    outDir: "dist",
    emptyOutDir: true
  }
});
