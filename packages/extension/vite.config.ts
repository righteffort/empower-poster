import { defineConfig } from "vite";

export default defineConfig({
  build: {
    modulePreload: { polyfill: false },
    rollupOptions: {
      input: {
        background: "src/background.ts",
        content: "src/content.ts",
        options: "src/options.html",
        onboarding: "src/onboarding.html",
      },
      output: {
        entryFileNames: "src/[name].js",
        chunkFileNames: "src/[name].js",
        assetFileNames: "assets/[name].[ext]",
      },
    },
  },
});
