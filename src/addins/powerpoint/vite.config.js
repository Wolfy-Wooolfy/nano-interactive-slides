import { defineConfig } from "vite";
import fs from "fs";
import path from "path";

const HOME = process.env.USERPROFILE || process.env.HOME;
const CERTS = path.join(HOME, ".office-addin-dev-certs");

export default defineConfig({
  server: {
    https: {
      key: fs.readFileSync(path.join(CERTS, "localhost.key")),
      cert: fs.readFileSync(path.join(CERTS, "localhost.crt")),
    },
    host: "localhost",
    port: 3000,
    strictPort: true
  },
  build: { outDir: "dist", emptyOutDir: true }
});
