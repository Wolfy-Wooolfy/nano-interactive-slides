import { defineConfig } from 'vite'
import fs from 'node:fs'
import path from 'node:path'

const certDir = path.join(process.env.USERPROFILE, '.office-addin-dev-certs')
const certPath = path.join(certDir, 'localhost.crt')
const keyPath  = path.join(certDir, 'localhost.key')

const rootDir = __dirname
const repoRoot = path.resolve(rootDir, '../../..')
const packagesDir = path.resolve(repoRoot, 'packages')

export default defineConfig({
  root: rootDir,
  server: {
    https: { cert: fs.readFileSync(certPath), key: fs.readFileSync(keyPath) },
    host: 'localhost',
    port: 3000,
    strictPort: true,
    fs: { allow: [rootDir, packagesDir, repoRoot] }
  },
  resolve: {
    alias: {
      '@sim': path.resolve(packagesDir, 'simulation-engine')
    }
  }
})
