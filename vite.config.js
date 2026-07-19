import { defineConfig } from 'vite';

export default defineConfig({
  build: {
    outDir: 'dist',
    rollupOptions: {
      output: {
        manualChunks: {
          xlsx: ['xlsx'],
          chart: ['chart.js'],
          supabase: ['@supabase/supabase-js'],
          emailjs: ['@emailjs/browser'],
        },
      },
    },
  },
});
