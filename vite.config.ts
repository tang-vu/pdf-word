import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react()],
  // Preserve function/class names for PDF.js and production fault acceptance.
  esbuild: { keepNames: true },
}) 