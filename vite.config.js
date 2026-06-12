import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  resolve: {
    dedupe: ['react', 'react-dom', '@supabase/supabase-js'],
  },
  optimizeDeps: {
    include: ['@supabase/supabase-js', 'xlsx'],
    esbuildOptions: { target: 'es2020' },
  },
  build: {
    target: 'es2020',
    sourcemap: 'inline',
    minify: false,
    commonjsOptions: { transformMixedEsModules: true },
    rollupOptions: {
      output: {
        manualChunks(id) {
          if (id.includes('@supabase')) return 'supabase';
          if (id.includes('xlsx')) return 'xlsx';
        }
      }
    }
  },
})
