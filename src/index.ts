// ── Types ──────────────────────────────────────────────────────────
export type {
  TextStyle,
  CellStyle,
  BorderStyle,
  TableStyle,
  FieldConfig,
  BasicSection,
  ListSection,
  SectionConfig,
  ExportOptions,
} from './types'

// ── Engine ─────────────────────────────────────────────────────────
export { buildDocument, exportWord } from './engine/builder'
export { downloadBlob, getValueByPath } from './engine/utils'
export {
  toTextRunProperties,
  toDocxAlignment,
  toDocxBorder,
  toDocxShading,
  defaultTitleStyle,
  defaultSectionTitleStyle,
  defaultHeaderCellStyle,
  defaultCellStyle,
  defaultLabelStyle,
} from './engine/styler'

// ── Composables ────────────────────────────────────────────────────
export { useWordExporter } from './composables/useWordExporter'

// ── Components ─────────────────────────────────────────────────────
export { default as WordExport } from './components/WordExport.vue'
