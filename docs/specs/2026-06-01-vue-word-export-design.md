# vue-word-export — Design Doc

> 基于 **Vue 3 + TypeScript + docx** 的 Word 文档导出组件。
> 用户根据后端返回的 JSON 数据，通过前端配置化方式生成 `.docx` 文件。

---

## 1. Problem Statement

前端需要将后端 API 返回的 JSON 数据导出为格式化的 Word (.docx) 文档。文档包含：

- **标题** — 文档主标题
- **基本信息区** — 键值对数据，渲染为两列表格（标签 | 值）
- **列表区** — 数组数据，渲染为多列表格（表头 + 数据行）

要求支持**配置化导出**：用户在前端通过类型安全的配置描述文档结构，决定哪些字段导出、如何排版。

## 2. Non-Goals

- 不做 Word 文档的**导入/解析**（只导出不导入）
- 不做富文本编辑器式的自由排版（按 Section 有序布局）
- 不做服务端渲染（纯浏览器端生成）
- 不依赖 Element Plus 或任何 UI 库（组件只提供 slot 供用户自定义触发按钮）

## 3. Architecture

### 3.1 Layered Design

严格分层，与已验证的 `vue-excel-export` 架构一致：

```
src/
├── types/           # 类型定义（纯 TS interface）
├── engine/          # 核心引擎（zero Vue dependency，可在 Node 测试）
├── composables/     # Vue 3 hooks
├── components/      # Vue SFC
└── index.ts         # 统一导出入口
```

### 3.2 依赖关系

```
types (pure TS)
  ↑
engine (pure TS, depends on docx)
  ↑
composables (depends on Vue 3 + engine)
  ↑
components (depends on Vue 3 + composables)
  ↑
consumer app (Vue 3 + peer deps: vue, docx)
```

- `engine/` 是纯 TypeScript，不含任何 Vue 导入
- `vue` 和 `docx` 是 peer dependencies，由消费者安装
- 库本身通过 Vite lib mode 打包

## 4. Type System

### 4.1 `TextStyle` — 文字样式

```typescript
interface TextStyle {
  bold?: boolean
  italic?: boolean
  fontSize?: number        // pt
  fontColor?: string       // hex: '#FF0000'
  fontName?: string        // '微软雅黑'
  underline?: boolean
}
```

### 4.2 `CellStyle` — 单元格样式

```typescript
interface CellStyle {
  bold?: boolean
  fontSize?: number
  fontColor?: string
  fontName?: string
  bgColor?: string
  align?: 'left' | 'center' | 'right'
  verticalAlign?: 'top' | 'middle' | 'bottom'
  border?: boolean | { style: 'single' | 'double' | 'dotted'; color?: string }
}
```

### 4.3 `TableStyle` — 表格样式

```typescript
interface TableStyle {
  headerStyle?: Partial<CellStyle>
  cellStyle?: Partial<CellStyle>
  border?: boolean           // 默认 true
  autoWidth?: boolean        // 默认 true
}
```

### 4.4 `FieldConfig` — 字段配置

描述一个字段在文档中的映射关系：

```typescript
interface FieldConfig {
  label: string                            // 显示标签（如 "企业名称"）
  field: string                            // 数据路径，支持点号嵌套（如 "company.name"）
  width?: number | string                  // 列宽（磅值如 3000，或百分比如 "20%"）
  align?: 'left' | 'center' | 'right'
  style?: Partial<CellStyle>
  format?: (value: unknown, row: Record<string, unknown>) => string  // 值格式化函数
}
```

### 4.5 `SectionConfig` — 文档段落配置

判别联合类型，engine 通过 `type` 字段选择渲染逻辑：

```typescript
/** 基本信息区 —— 键值对表格 */
interface BasicSection {
  type: 'basic'
  title?: string                  // 区域小标题（可选）
  fields: FieldConfig[]           // 按序显示
  columns?: 1 | 2                 // 表格列数：1=单列(label|value)，2=双列(label|value + label|value)
  style?: Partial<TableStyle>
}

/** 列表区 —— 数据数组表格 */
interface ListSection {
  type: 'list'
  title?: string                  // 列表标题（可选）
  dataField: string               // 数组数据路径
  columns: FieldConfig[]          // 列配置
  style?: Partial<TableStyle>
}

type SectionConfig = BasicSection | ListSection
```

### 4.6 `ExportOptions` — 导出选项

```typescript
interface ExportOptions {
  title?: {
    text: string | (() => string)       // 文档标题文字
    style?: Partial<TextStyle>
  }
  sections: SectionConfig[]              // 有序的文档段落
  data: Record<string, unknown>          // 后端 JSON 数据
  filename?: string                      // 默认 'export.docx'
  pageOrientation?: 'portrait' | 'landscape'
  margins?: {
    top?: number
    bottom?: number
    left?: number
    right?: number
  }
}
```

## 5. Engine Layer

### 5.1 文件结构

```
src/engine/
├── builder.ts       # buildDocument() — 编排所有渲染逻辑，返回 Blob
├── table.ts         # buildBasicTable(), buildListTable() — 表格构建
├── styler.ts        # 类型映射：our styles → docx IRunProperties/ITableCellProperties
└── utils.ts         # 工具函数：getValueByPath(), downloadBlob()
```

### 5.2 `builder.ts` — 文档构建编排

```
ExportOptions
  │
  ▼
Document
  ├── Section 1:
  │   ├── Title (Paragraph, HeadingLevel.TITLE)
  │   ├── BasicSection → buildBasicTable()
  │   │    2列: [label(加粗), value] 逐行
  │   │    或 1列: label(加粗): value 逐行
  │   ├── (空段落间隔)
  │   └── ListSection → buildListTable()
  │        [header1, header2, ...]
  │        [row1val1, row1val2, ...]
  │        [row2val1, row2val2, ...]
  │
  └── (Multiple sections supported)
  │
  ▼
Packer.toBlob(doc) → Blob → download
```

核心函数签名：

```typescript
export async function buildDocument(options: ExportOptions): Promise<Blob>
```

### 5.3 `table.ts` — 表格构建

**基本信息表 (BasicSection)**:

- `columns: 2`（默认）：每行两个单元格，左侧 label（加粗）| 右侧 value
- `columns: 1`：每行一个单元格，内容为 `label: value`（label 部分加粗）
- 区域标题（如有）作为单独段落渲染在表格上方

**列表表 (ListSection)**:

- 第一行为表头（thead），内容为各列 `FieldConfig.label`
- 后续为数据行，遍历 `data[dataField]` 数组
- 每列通过 `field` 路径取值，通过 `format` 函数格式化（如有）
- 支持 `align` 控制列对齐

### 5.4 `styler.ts` — 样式映射

| Our type | docx API |
|----------|----------|
| `TextStyle.fontSize` (pt) | `{ size: fontSize * 2 }` (half-points) |
| `TextStyle.fontColor` (`#RRGGBB`) | `{ color: RRGGBB }` (strip `#`) |
| `TextStyle.bold` | `{ bold: true }` |
| `CellStyle.align` | `AlignmentType.CENTER / LEFT / RIGHT` |
| `CellStyle.bgColor` | `{ shading: { fill: RRGGBB } }` |
| `CellStyle.border` | `BorderStyle.SINGLE / DOUBLE / DOTTED` |

### 5.5 `utils.ts` — 工具函数

```typescript
/** 从嵌套对象中按点号路径取值 */
export function getValueByPath(obj: Record<string, unknown>, path: string): unknown

/** 触发浏览器下载 Blob */
export function downloadBlob(blob: Blob, filename: string): void
```

`getValueByPath` 实现：

```typescript
function getValueByPath(obj: unknown, path: string): unknown {
  return path.split('.').reduce((acc, key) => {
    if (acc && typeof acc === 'object') return (acc as Record<string, unknown>)[key]
    return undefined
  }, obj)
}
```

支持如 `"user.address.city"` 的嵌套字段路径。

## 6. Composable Layer

### `useWordExporter()`

```typescript
export function useWordExporter() {
  const loading = ref(false)
  const error = shallowRef<Error | null>(null)

  /** 构建并下载 Word 文档 */
  async function exportWord(options: ExportOptions): Promise<void>

  /** 仅构建文档，返回 Blob（不触发下载） */
  async function build(options: ExportOptions): Promise<Blob>

  return { exportWord, build, loading, error }
}
```

## 7. Component Layer

### `<WordExport>` Props

| Prop | Type | Default | Description |
|------|------|---------|-------------|
| `data` | `Record<string, unknown>` | — | JSON 数据（必填） |
| `sections` | `SectionConfig[]` | — | 文档 Section 配置（必填） |
| `title` | `{ text: string, style? }` | — | 文档标题 |
| `filename` | `string` | `'export.docx'` | 导出文件名 |
| `pageOrientation` | `'portrait' \| 'landscape'` | `'portrait'` | 页面方向 |
| `margins` | `{ top?, bottom?, left?, right? }` | — | 页边距（twips） |

### Events

| Name | Payload | Description |
|------|---------|-------------|
| `before-export` | — | 导出前触发 |
| `after-export` | — | 导出成功 |
| `error` | `Error` | 导出失败 |

### Slots

| Name | Scope | Description |
|------|-------|-------------|
| `trigger` | `{ exportWord: () => void, loading: boolean }` | 自定义触发按钮 |

## 8. Word 文档效果

```
┌─────────────────────────────────────────┐
│           企业信用报告                    │  ← Title
│                                         │
│  基本信息                               │  ← BasicSection title
│  ┌──────────────┬──────────────────────┐│
│  │ 企业名称     │ 某某科技有限公司      ││
│  │ 统一信用代码 │ 91110000MA12345678   ││  ← 2列基本信息表
│  │ 法定代表人   │ 张三                  ││
│  │ 注册资本    │ 1000 万元             ││
│  └──────────────┴──────────────────────┘│
│                                         │
│  股东信息                               │  ← ListSection title
│  ┌────────────┬──────────┬────────────┐│
│  │ 股东名称   │ 出资比例  │ 出资额      ││  ← 表头行
│  ├────────────┼──────────┼────────────┤│
│  │ 张三       │ 60.0%    │ 600        ││
│  │ 李四       │ 40.0%    │ 400        ││  ← 数据行
│  └────────────┴──────────┴────────────┘│
└─────────────────────────────────────────┘
```

## 9. Tech Stack

| Component | Choice |
|-----------|--------|
| Framework | Vue 3 (Composition API) |
| Language | TypeScript (strict mode) |
| Word engine | `docx` (dolanmiu/docx) |
| Build | Vite 6+ (lib mode) |
| Test | Vitest 3+ |
| Type check | vue-tsc |
| Peer deps | `vue ^3.5`, `docx ^8` |

## 10. Package Structure

```json
{
  "name": "vue-word-export",
  "type": "module",
  "main": "./dist/vue-word-export.umd.cjs",
  "module": "./dist/vue-word-export.js",
  "types": "./dist/index.d.ts",
  "files": ["dist"],
  "scripts": {
    "dev": "vite",
    "build": "vue-tsc --noEmit && vite build",
    "test": "vitest run",
    "typecheck": "vue-tsc --noEmit"
  },
  "peerDependencies": {
    "vue": "^3.5.0",
    "docx": "^8.0.0"
  }
}
```

## 11. Build Configuration

Vite lib mode 配置（参考 vue-excel-export）：

- 入口：`src/index.ts`
- External：`vue`, `docx`
- Plugin：`@vitejs/plugin-vue`, `vite-plugin-dts`
- 输出：ESM + UMD + .d.ts

## 12. Testing Strategy

- **Engine 层**（纯 TS）：Vitest + Node 环境
  - `builder.test.ts` — 验证 `buildDocument()` 返回合法 Blob
  - `table.test.ts` — 验证 BasicSection / ListSection 的表格结构
  - `utils.test.ts` — 验证 `getValueByPath` 路径解析
- 不测试 docx 输出的 XML 内容细节（docx 库保证格式正确性）

## 13. Self-Review Check

- [x] **Placeholder scan**: 无 TBD/TODO/不完整段落
- [x] **Internal consistency**: 类型定义 ↔ Engine 实现 ↔ Composable ↔ Component，每层接口一致
- [x] **Scope check**: 聚焦于"标题 + 基本信息表 + 列表表"的 Section 配置模式，适合单次实现
- [x] **Ambiguity check**: 
  - `columns?: 1 | 2` 的默认值明确为 2
  - `field` 支持点号路径，行为已定义
  - `format` 函数的签名和类型已明确
  - Engine 层纯 TS，不依赖 Vue，边界清晰
