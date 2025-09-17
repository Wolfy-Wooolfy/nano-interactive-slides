import { defineConfig } from 'vite'
import fs from 'node:fs'
import path from 'node:path'

const certDir = path.join(process.env.USERPROFILE, '.office-addin-dev-certs')
const certPath = path.join(certDir, 'localhost.crt')
const keyPath  = path.join(certDir, 'localhost.key')

export default defineConfig({
  server: {
    https: {
      cert: fs.readFileSync(certPath),
      key: fs.readFileSync(keyPath)
    },
    host: 'localhost',
    port: 3000,
    strictPort: true
  }
})
