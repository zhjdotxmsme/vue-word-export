import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import dts from 'vite-plugin-dts'
import { resolve } from 'path'

export default defineConfig({
  plugins: [
    vue(),
    dts({ tsconfigPath: './tsconfig.json', include: ['src/**/*.ts', 'src/**/*.vue'] }),
  ],
  build: {
    lib: {
      entry: resolve(__dirname, 'src/index.ts'),
      name: 'VueWordExport',
      fileName: 'vue-word-export',
    },
    rollupOptions: {
      external: ['vue', 'docx', 'echarts'],
      output: {
        globals: {
          vue: 'Vue',
          docx: 'docx',
        },
      },
    },
  },
  resolve: {
    alias: {
      '@': resolve(__dirname, 'src'),
    },
  },
})
