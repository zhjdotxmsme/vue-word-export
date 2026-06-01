# vue-word-export

基于 **Vue 3 + TypeScript + [docx](https://github.com/dolanmiu/docx)** 的 Word 文档导出组件。将后端 JSON 数据渲染为格式化的 `.docx` 文件，支持基本信息区和列表数据区，全配置化导出。

- 🚀 **零框架依赖** — engine 层纯 TypeScript，无 Vue 耦合
- 📦 **轻量** — 打包后 ~3.2 KB (gzip)
- 🧩 **双形态** — Hook (`useWordExporter`) + 组件 (`<WordExport>`)
- 📝 **基本信息区** — 无边框多列并排，内容超长自动列内换行
- 📊 **列表数据区** — 标准表格（表头 + 数据行），支持格式化
- 🎨 **自定义样式** — 字体/颜色/对齐
- 🌳 **Tree-shakable** — 按需导入

---

## 安装

```bash
npm install vue-word-export docx
```

> `docx` 是 peer dependency，需要单独安装。

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

Word 文档被定义为有序的 Section 列表，每个 Section 可以是 `basic` 或 `list`：

```typescript
type SectionConfig = BasicSection | ListSection
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

---

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

# 构建
npm run build

# 启动示例项目
cd example && npm install && npm run dev
```

---

## 技术栈

- **Vue 3** — Composition API + TypeScript
- **docx** ^8.0.0 — Word 文档生成库（[GitHub](https://github.com/dolanmiu/docx)）
- **Vite** — 库模式构建
- **Vitest** — 单元测试

## License

MIT
