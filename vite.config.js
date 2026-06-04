import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// Served at domain root on Vercel/Netlify, so base is "/".
export default defineConfig({
  base: "/",
  plugins: [react()],
});
