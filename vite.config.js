import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// `base` must match the repo name for GitHub Pages project sites
// (served from https://<user>.github.io/qs-workspace/).
export default defineConfig({
  base: "/qs-workspace/",
  plugins: [react()],
});
