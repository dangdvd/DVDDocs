import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// base './' để bản build chạy được trong Electron (file://) và Capacitor.
export default defineConfig({
  base: './',
  plugins: [react()],
  build: { outDir: 'dist', chunkSizeWarningLimit: 4000 },
  test: { environment: 'node' },
});
