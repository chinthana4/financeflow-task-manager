import { defineConfig } from 'vite';

export default defineConfig({
  test: {
    environment: 'jsdom',
    // jsdom defaults to about:blank, which has no http(s) scheme — unlike a
    // real browser, which always runs on an actual origin. sanitizeUrl()
    // resolves scheme-less input relative to window.location, so tests need
    // a realistic origin here to match real deployed behavior.
    environmentOptions: { jsdom: { url: 'https://financeflow-taskmanager.onrender.com/' } },
    globals: false,
  },
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
