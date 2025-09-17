import { defineConfig } from 'vite'
import fs from 'fs'
import path from 'path'

const certDir = path.join(process.env.USERPROFILE, '.office-addin-dev-certs')

export default defineConfig({
  server: {
    https: {
      key: fs.readFileSync(path.join(certDir, 'localhost.key')),
      cert: fs.readFileSync(path.join(certDir, 'localhost.crt'))
    },
    port: 3000,
    strictPort: true,
    host: 'localhost'
  },
  preview: {
    https: {
      key: fs.readFileSync(path.join(certDir, 'localhost.key')),
      cert: fs.readFileSync(path.join(certDir, 'localhost.crt'))
    },
    port: 3000,
    strictPort: true,
    host: 'localhost'
  }
})
