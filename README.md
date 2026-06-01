# vue-word-export

基于 **Vue 3 + TypeScript + [docx](https://github.com/dolanmiu/docx)** 的 Word 文档导出组件。将后端 JSON 数据渲染为格式化的 `.docx` 文件，支持多种文档段落类型，全配置化导出。

- 🚀 **零框架依赖** — engine 层纯 TypeScript，无 Vue 耦合
- 📦 **轻量** — 打包后 ~3.5 KB (gzip)
- 🧩 **双形态** — Hook (`useWordExporter`) + 组件 (`<WordExport>`)
- 📝 **基本信息区** — 无边框多列并排，内容超长自动列内换行
- 📊 **列表数据区** — 标准表格（表头 + 数据行），支持格式化
- 🎨 **封面页** — 自定义标题/副标题/日期/Logo，支持背景色和居中布局
- 📄 **富文本段落** — 纯文本 + 混排样式（段内加粗/颜色/字体混排）
- 📈 **图表** — 折线图/柱状图/饼图/面积图（基于 ECharts SVG 渲染）
- 📌 **页眉/页脚/页码** — 左/中/右三区，支持页码、总页数、日期
- 🎨 **全局主题** — 统一默认样式，所有值可覆盖
- 🌳 **Tree-shakable** — 按需导入

---

## 安装

```bash
npm install vue-word-export docx
```

> `docx` 是 peer dependency，需要单独安装。
> `echarts` 是可选依赖（仅在 ChartSection 使用时需要）：`npm install echarts`

---

## 快速开始

### 方式一：组件式

```vue
<script setup lang="ts">
import { WordExport } from 'vue-word-export'
import type { SectionConfig } from 'vue-word-export'

const sections: SectionConfig[] = [
  {
    type: 'basic',
    title: '基本信息',
    fields: [
      { label: '企业名称', field: 'companyName' },
      { label: '法定代表人', field: 'legalPerson' },
    ],
  },
  {
    type: 'list',
    title: '股东信息',
    dataField: 'shareholders',
    columns: [
      { label: '股东名称', field: 'name' },
      { label: '出资比例', field: 'ratio',
        format: (v) => `${(Number(v) * 100).toFixed(1)}%` },
    ],
  },
]

const data = {
  companyName: '某某科技有限公司',
  legalPerson: '张三',
  shareholders: [
    { name: '张三', ratio: 0.6 },
    { name: '李四', ratio: 0.3 },
  ],
}
</script>

<template>
  <WordExport
    :data="data"
    :sections="sections"
    title="企业信用报告"
    filename="报告.docx"
  >
    <template #trigger="{ exportWord, loading }">
      <button :disabled="loading" @click="exportWord">
        {{ loading ? '生成中...' : '导出报告' }}
      </button>
    </template>
  </WordExport>
</template>
```

### 方式二：Hook 式

```typescript
import { useWordExporter } from 'vue-word-export'
import type { SectionConfig } from 'vue-word-export'

const { exportWord, loading } = useWordExporter()

async function handleExport() {
  await exportWord({
    filename: '报告.docx',
    title: '项目进度报告',
    sections: [
      {
        type: 'basic',
        fields: [
          { label: '项目名称', field: 'name' },
          { label: '状态', field: 'status' },
        ],
      },
      {
        type: 'list',
        dataField: 'tasks',
        columns: [
          { label: '任务', field: 'taskName' },
          { label: '完成度', field: 'progress',
            format: (v) => `${v}%` },
        ],
      },
    ],
    data: {
      name: '系统重构',
      status: '进行中',
      tasks: [
        { taskName: '需求分析', progress: 100 },
        { taskName: '前端开发', progress: 60 },
      ],
    },
  })
}
```

---

## 核心概念

### SectionConfig — 文档段落配置

Word 文档被定义为有序的 Section 列表，目前支持 **5 种 Section 类型**：

```typescript
type SectionConfig =
  | BasicSection        // 基本信息区（键值对）
  | ListSection         // 列表数据区（表格）
  | CoverSection        // 封面页
  | ParagraphSection    // 富文本段落
  | ChartSection        // 图表
```

### BasicSection — 基本信息区

渲染为**无边框多列表格**，视觉上与段落排版无异，内容超长时自动在列内换行。

```typescript
interface BasicSection {
  type: 'basic'
  title?: string                  // 区域标题（可选）
  fields: FieldConfig[]           // 字段映射列表
  columns?: 1 | 2                 // 每行几列，默认 2
  style?: Partial<TableStyle>
}
```

### ListSection — 列表数据区

渲染为**标准表格**（有边框），第一行为表头，后续为数据行。

```typescript
interface ListSection {
  type: 'list'
  title?: string                  // 列表标题（可选）
  dataField: string               // 数组数据字段路径
  columns: FieldConfig[]           // 列配置
  style?: Partial<TableStyle>
}
```

### CoverSection — 封面页

渲染为独立封面页，自动分页。支持标题、副标题、日期、正文、Logo 等多种元素。

```typescript
interface CoverSection {
  type: 'cover'
  items: CoverItem[]               // 封面元素列表（按序渲染）
  style?: CoverStyle
}

// 支持的元素类型
type CoverItem = CoverTitle | CoverSubtitle | CoverDate | CoverText | CoverImage

interface CoverTitle {
  type: 'title'
  text: string | (() => string)
  style?: Partial<TextStyle>
}

interface CoverImage {
  type: 'image'
  src: string | (() => string | Promise<string>)
  width?: number
  height?: number
  align?: 'left' | 'center' | 'right'
}
```

**示例：**
```typescript
{
  type: 'cover',
  items: [
    { type: 'image', src: logoBase64, width: 200 },
    { type: 'title', text: '企业信用报告',
      style: { fontSize: 36, bold: true } },
    { type: 'subtitle', text: 'Enterprise Credit Report' },
    { type: 'date', format: 'YYYY年MM月DD日' },
  ]
}
```

### ParagraphSection — 富文本段落

支持纯文本段落和段内样式混排。

```typescript
interface ParagraphSection {
  type: 'paragraph'
  title?: string                                  // 可选标题
  content: TextRunFragment | TextRunFragment[]    // 文本块或混合文本块数组
  style?: Partial<ParagraphStyle>
}

interface TextRunFragment {
  type: 'text'
  text: string | (() => string)
  style?: Partial<TextStyle>
}
```

**纯文本段落：**
```typescript
{
  type: 'paragraph',
  title: '项目背景',
  content: { type: 'text', text: '本项目旨在建设一套企业级数据管理平台。' },
  style: { indentFirstLine: 480, lineSpacing: 360 }
}
```

**富文本混排：**
```typescript
{
  type: 'paragraph',
  content: [
    { type: 'text', text: '经审核，该企业' },
    { type: 'text', text: '符合', style: { bold: true, fontColor: '#4CAF50' } },
    { type: 'text', text: '本次申报条件。' },
  ]
}
```

### ChartSection — 图表

渲染为静态图片嵌入文档。**需要安装 `echarts` 作为可选依赖。**

```typescript
interface ChartSection {
  type: 'chart'
  title?: string
  chartType: 'line' | 'bar' | 'pie' | 'area' | 'scatter'
  dataField: string              // 数组数据路径
  dimension: ChartDimension
  width?: number                 // 图片宽度（dxa）
  height?: number                // 图片高度（dxa）
  style?: ChartStyle
}

interface ChartDimension {
  categoryField: string          // X轴/饼图标签字段
  valueFields: ChartValueField[] // 数值字段列表（支持多系列）
}
```

**示例：**
```typescript
{
  type: 'chart',
  title: '月度营收趋势',
  chartType: 'line',
  dataField: 'monthlyTrend',
  dimension: {
    categoryField: 'month',
    valueFields: [{ field: 'revenue', name: '营收' }]
  }
}
```

### FieldConfig — 字段配置

```typescript
interface FieldConfig {
  label: string                    // 显示标签（如 "企业名称"）
  field: string                    // 数据路径（支持点号嵌套）
  width?: number | string          // 列宽
  align?: 'left' | 'center' | 'right'
  style?: Partial<CellStyle>
  format?: (value: unknown, row: Record<string, unknown>) => string  // 格式化
}
```

### 多层数据结构

`field` 支持点号路径，方便从嵌套 JSON 中取值：

```typescript
const data = {
  user: { name: '张三', contact: { phone: '13800138000' } },
}

const field = { label: '手机号', field: 'user.contact.phone' }
// → 取到 '13800138000'
```

### 值格式化

通过 `format` 函数自定义值的显示方式：

```typescript
{ label: '出资比例', field: 'ratio',
  format: (v) => `${(Number(v) * 100).toFixed(1)}%` }

{ label: '状态', field: 'status',
  format: (v) => {
    const map = { done: '已完成', doing: '进行中', todo: '待开始' }
    return map[String(v)] || String(v)
  }
}

{ label: '金额', field: 'amount', align: 'right',
  format: (v) => `¥${Number(v).toLocaleString()}` }
```

---

## 文档输出效果

```
┌──────────────────────────────────────────┐
│           企业信用报告                     │  ← 标题
│                                          │
│  基本信息                                │  ← 区域标题
│  企业名称：某某科技有限公司    统一信用代码：91110000MA12345678
│  法定代表人：张三                注册资本：1000 万元  ← 无边框两列
│  经营范围：计算机软硬件的        企业状态：存续
│  技术开发、技术咨询
│  ↑ 内容超长时自动列内换行
│                                          │
│  股东信息                                │  ← 列表标题
│  ┌──────────┬──────────┬──────────────┐  │
│  │ 股东名称 │ 出资比例 │ 认缴出资额   │  │  ← 有边框表格
│  ├──────────┼──────────┼──────────────┤  │
│  │ 张三     │ 60.0%    │ 600          │  │
│  │ 李四     │ 30.0%    │ 300          │  │
│  └──────────┴──────────┴──────────────┘  │
└──────────────────────────────────────────┘
```

---

## API 参考

### useWordExporter()

```typescript
function useWordExporter(): {
  exportWord: (options: ExportOptions) => Promise<void>
  build: (options: ExportOptions) => Promise<Blob>
  loading: Ref<boolean>
  error: Ref<Error | null>
}
```

| 返回值 | 类型 | 说明 |
|--------|------|------|
| `exportWord` | `(options) => Promise<void>` | 构建 + 下载一步完成 |
| `build` | `(options) => Promise<Blob>` | 仅构建文档，返回 Blob |
| `loading` | `Ref<boolean>` | 导出进行中 |
| `error` | `Ref<Error \| null>` | 导出异常 |

### ExportOptions

```typescript
interface ExportOptions {
  title?: string | {                // 文档标题
    text: string | (() => string)
    style?: Partial<TextStyle>
  }
  sections: SectionConfig[]         // 文档段落配置（必填）
  data: Record<string, unknown>     // JSON 数据（必填）
  filename?: string                 // 文件名，默认 'export.docx'
  pageOrientation?: 'portrait' | 'landscape'
  margins?: { top? bottom? left? right? }
  header?: HeaderFooterConfig       // 页眉配置
  footer?: HeaderFooterConfig       // 页脚配置
  differentFirstPage?: boolean      // 首页不同（封面页不显示页眉页脚）
  theme?: DocumentTheme             // 全局样式主题
}
```

### WordExport 组件 Props

| 名称 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `data` | `Record<string, unknown>` | — | JSON 数据（必填） |
| `sections` | `SectionConfig[]` | — | 文档配置（必填） |
| `title` | `string \| { text, style }` | — | 文档标题 |
| `filename` | `string` | `'export.docx'` | 导出文件名 |
| `pageOrientation` | `'portrait' \| 'landscape'` | `'portrait'` | 页面方向 |
| `margins` | `{ top?, bottom?, left?, right? }` | — | 页边距 |

### 页眉/页脚

```typescript
interface HeaderFooterConfig {
  left?: HeaderFooterContent        // 左侧内容
  center?: HeaderFooterContent      // 中间内容
  right?: HeaderFooterContent       // 右侧内容
  border?: boolean                  // 分割线（默认 true）
}

type HeaderFooterContent =
  | string
  | { text: string | (() => string); style?: Partial<TextStyle> }
  | ({ text: string | (() => string); style?: Partial<TextStyle> } | { type: 'pageNumber' } | { type: 'totalPages' } | { type: 'date'; format?: string } | { type: 'image'; src: string; width?: number; height?: number })[]
```

**示例：**
```typescript
{
  header: {
    center: { text: '某某科技有限公司', style: { fontSize: 9, fontColor: '#999' } },
  },
  footer: {
    left: { type: 'date', format: 'YYYY-MM-DD' },
    center: { type: 'pageNumber' },
    right: [
      { text: '第 ' },
      { type: 'pageNumber' },
      { text: ' 页 / 共 ' },
      { type: 'totalPages' },
      { text: ' 页' },
    ],
  },
  differentFirstPage: true,
}
```

### WordExport 组件 Events

| 名称 | 参数 | 说明 |
|------|------|------|
| `before-export` | — | 导出前触发 |
| `after-export` | — | 导出后触发 |
| `error` | `Error` | 导出失败 |

### WordExport 组件 Slots

| 名称 | 参数 | 说明 |
|------|------|------|
| `trigger` | `{ exportWord: () => void, loading: boolean }` | 自定义触发按钮 |

---

## 样式

### TextStyle

```typescript
interface TextStyle {
  bold?: boolean
  italic?: boolean
  fontSize?: number       // pt
  fontColor?: string      // hex: '#FF0000'
  fontName?: string       // '微软雅黑'
  underline?: boolean
}
```

### CellStyle

```typescript
interface CellStyle {
  bold?: boolean
  fontSize?: number
  fontColor?: string
  fontName?: string
  bgColor?: string
  align?: 'left' | 'center' | 'right'
  verticalAlign?: 'top' | 'middle' | 'bottom'
  border?: boolean | { style: 'single' | 'double' | 'dotted', color?: string }
}
```

### ParagraphStyle — 段落样式

```typescript
interface ParagraphStyle {
  align?: 'left' | 'center' | 'right' | 'justify'
  spacingBefore?: number    // 段前间距（twips）
  spacingAfter?: number     // 段后间距（twips）
  lineSpacing?: number      // 行距
  indentFirstLine?: number  // 首行缩进（twips）
  indentLeft?: number
  indentRight?: number
  headingLevel?: 1 | 2 | 3 | 4 | 5 | 6
  pageBreakBefore?: boolean
  keepWithNext?: boolean
  keepLines?: boolean
}
```

### ChartStyle — 图表样式

```typescript
interface ChartStyle {
  colorPalette?: string[]   // 调色板
  showLegend?: boolean
  legendPosition?: 'top' | 'bottom' | 'left' | 'right'
  xAxisName?: string
  yAxisName?: string
  showDataLabel?: boolean
  backgroundColor?: string
}
```

### CoverStyle — 封面样式

```typescript
interface CoverStyle {
  backgroundColor?: string
  verticalAlign?: 'middle' | 'top' | 'bottom'
  itemSpacing?: number              // 元素间距（twips）
  margins?: { top?, bottom?, left?, right? }
}
```

### DocumentTheme — 全局主题

通过 `theme` 设置全局默认样式，所有值可被各 Section 的 `style` 覆盖。

```typescript
interface DocumentTheme {
  fontFamily?: string              // 默认 '微软雅黑'
  fontSize?: number                // 默认 10.5pt
  title?: { style?: Partial<TextStyle> }
  cover?: Partial<CoverStyle>
  paragraph?: {
    style?: Partial<ParagraphStyle>
    heading?: Record<number, Partial<TextStyle>>
  }
  table?: {
    basic?: { labelStyle?, valueStyle?, border? }
    list?: { headerStyle?, cellStyle?, border? }
  }
  chart?: { colorPalette?: string[]; style?: Partial<ChartStyle> }
  headerFooter?: { style?: Partial<TextStyle>; border?: boolean }
}
```

**样式优先级：** `字段/style` > `section/style` > `theme/对应类型` > `内置默认值`

## 完整示例

### 企业信用报告

```vue
<script setup lang="ts">
import { WordExport } from 'vue-word-export'
import type { SectionConfig, FieldConfig } from 'vue-word-export'

const sections: SectionConfig[] = [
  {
    type: 'basic',
    title: '基本信息',
    fields: [
      { label: '企业名称', field: 'companyName' },
      { label: '统一社会信用代码', field: 'creditCode' },
      { label: '法定代表人', field: 'legalPerson' },
      { label: '注册资本', field: 'registeredCapital',
        format: (v) => `${v} 万元` },
      { label: '成立日期', field: 'establishDate' },
      { label: '企业状态', field: 'status' },
    ],
  },
  {
    type: 'list',
    title: '股东信息',
    dataField: 'shareholders',
    columns: [
      { label: '股东名称', field: 'name' },
      { label: '出资比例', field: 'ratio',
        format: (v) => `${(Number(v) * 100).toFixed(1)}%` },
      { label: '认缴出资额（万元）', field: 'amount', align: 'right' },
    ],
  },
]

const data = {
  companyName: '某某科技有限公司',
  creditCode: '91110000MA12345678',
  legalPerson: '张三',
  registeredCapital: 1000,
  establishDate: '2015-03-15',
  status: '存续',
  shareholders: [
    { name: '张三', ratio: 0.6, amount: 600 },
    { name: '李四', ratio: 0.3, amount: 300 },
  ],
}
</script>
```

### 使用 Hook 封装业务

```typescript
// composables/useReportExport.ts
import { useWordExporter } from 'vue-word-export'
import type { SectionConfig } from 'vue-word-export'

export function useReportExport() {
  const { exportWord, loading } = useWordExporter()

  const reportSections: SectionConfig[] = [
    {
      type: 'basic',
      title: '项目概览',
      fields: [
        { label: '项目名称', field: 'name' },
        { label: '项目经理', field: 'manager' },
        { label: '总体进度', field: 'progress',
          format: (v) => `${v}%` },
      ],
    },
    {
      type: 'list',
      title: '任务明细',
      dataField: 'tasks',
      columns: [
        { label: '任务', field: 'name' },
        { label: '负责人', field: 'assignee' },
        { label: '状态', field: 'status' },
        { label: '完成度', field: 'progress',
          format: (v) => `${v}%`, align: 'right' },
      ],
    },
  ]

  async function exportReport(data: Record<string, unknown>) {
    return exportWord({
      title: '项目进度报告',
      sections: reportSections,
      data,
      filename: `项目报告_${new Date().toISOString().slice(0, 10)}.docx`,
    })
  }

  return { exportReport, loading }
}
```

---

## 开发

```bash
# 安装依赖
cd vue-word-export && npm install

# 类型检查
npm run typecheck

# 测试
npm run test

# 构建
npm run build

# 启动示例项目
cd example && npm install && npm run dev
```

---

## 技术栈

- **Vue 3** — Composition API + TypeScript
- **docx** ^8.0.0 — Word 文档生成库（[GitHub](https://github.com/dolanmiu/docx)）
- **echarts** — 图表渲染（可选，ChartSection 需要）
- **Vite** — 库模式构建
- **Vitest** — 单元测试

## License

MIT
