import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'

export default defineConfig({
  plugins: [vue()],
  server: {
    fs: {
      allow: ['..'],
    },
  },
  optimizeDeps: {
    exclude: ['vue-word-export'],
  },
  build: {
    rollupOptions: {
      external: ['echarts'],
    },
  },
})
