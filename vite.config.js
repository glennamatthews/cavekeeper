import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  optimizeDeps: {
    include: [
      'recharts',
      'recharts/es6/component/DefaultLegendContent',
      '@supabase/supabase-js',
      'xlsx',
    ],
    force: true,
  },
  build: {
    commonjsOptions: {
      transformMixedEsModules: true,
      include: [/recharts/, /node_modules/],
    },
    rollupOptions: {
      output: {
        manualChunks: {
          recharts: ['recharts'],
          supabase: ['@supabase/supabase-js'],
          xlsx: ['xlsx'],
        }
      }
    }
  },
})
