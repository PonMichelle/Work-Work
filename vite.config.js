import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// Relative base ("./") so assets load whether served from a domain root
// (Firebase Hosting: qs-workspace-5ddbd.web.app) or a sub-path (GitHub Pages).
export default defineConfig({
  base: "./",
  plugins: [react()],
});
