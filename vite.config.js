import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  server: {
    port: 5173,
    open: true,        // פותח את הדפדפן לבד
  },
  build: {
    outDir: "dist",
  },
});
