import { defineConfig } from "vite";
import { resolve } from "path";

// Dedicated build for the run-script CLI. The library build (vite.config.ts)
// targets the browser/neutral platform and is intentionally kept free of Node
// built-ins and of the WSF parser's fast-xml-parser dependency. This config
// bundles the CLI (and the host objects it pulls in) into a self-contained
// Node ESM executable, externalizing only Node built-ins.
export default defineConfig({
  build: {
    lib: {
      entry: resolve(__dirname, "src/cli/run-script.ts"),
      formats: ["es"],
      fileName: "run-script",
    },
    rollupOptions: {
      external: [/^node:/],
    },
    outDir: "dist",
    emptyOutDir: false,
    target: "node18",
    sourcemap: false,
    minify: false,
  },
});
