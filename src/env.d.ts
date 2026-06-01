/// <reference types="vite/client" />

declare module '*.vue' {
  import type { DefineComponent } from 'vue'
  const component: DefineComponent<{}, {}, any>
  export default component
}

// echarts 是 ChartSection 的可选依赖，动态导入时无需类型检查
declare module 'echarts' {
  const echarts: any
  export default echarts
}
