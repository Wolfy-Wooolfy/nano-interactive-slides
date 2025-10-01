import { defineConfig } from "vite";
import fs from "fs";
import path from "path";

const HOME = process.env.USERPROFILE || process.env.HOME;
const CERTS_DIR = path.join(HOME, ".office-addin-dev-certs");
const keyPath = path.join(CERTS_DIR, "localhost.key");
const certPath = path.join(CERTS_DIR, "localhost.crt");
const hasCerts = fs.existsSync(keyPath) && fs.existsSync(certPath);

export default defineConfig({
  server: {
    port: 3000,           // مهم: نفس البورت اللي في manifest
    strictPort: true,     // لو البورت مش فاضي يفشل بدل ما يبدّل لبورت تاني
    https: hasCerts
      ? {
          key: fs.readFileSync(keyPath),
          cert: fs.readFileSync(certPath),
        }
      : true,
    proxy: {
      "/nano": {
        target: "http://localhost:8787",
        changeOrigin: true,
        secure: false,
      },
    },
    hmr: { overlay: false }
  },
  build: { outDir: "dist", emptyOutDir: true },
});
