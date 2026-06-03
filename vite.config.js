import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  server: {
    port: 3000,
    proxy: {
      // Em dev, redireciona /api para o Express na porta 3001
      "/api": "http://localhost:3001",
    },
  },
});
