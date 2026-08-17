import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  define: {
    "process.env.NODE_ENV": JSON.stringify("production"),
  },
  build: {
    outDir: "dist",
    emptyOutDir: true,
    sourcemap: false,
    lib: {
      entry: "src/main.jsx",
      name: "ProjectPagesEditor",
      formats: ["iife"],
      fileName: () => "project-pages-editor.js",
      cssFileName: "project-pages-editor",
    },
    rollupOptions: {
      output: {
        assetFileNames: (assetInfo) =>
          assetInfo.name === "style.css"
            ? "project-pages-editor.css"
            : "project-pages-editor-[name][extname]",
      },
    },
  },
});
