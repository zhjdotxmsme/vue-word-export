// ── Types ──────────────────────────────────────────────────────────
export type {
  TextStyle, CellStyle, BorderStyle, TableStyle,
  ParagraphStyle, CoverStyle, ChartStyle, ChartDimension, ChartValueField, DocumentTheme,
  FieldConfig, BasicSection, ListSection, SectionConfig,
  CoverSection, CoverItem, CoverTitle, CoverSubtitle, CoverDate, CoverText, CoverImage,
  ParagraphSection, TextRunFragment, BreakFragment, ParagraphContent,
  ChartSection,
  ExportOptions,
  HeaderFooterConfig, HeaderFooterContent, HeaderFooterElement, HeaderFooterText,
} from './types'

// ── Engine ─────────────────────────────────────────────────────────
export { buildDocument, exportWord } from './engine/builder'
export { downloadBlob, getValueByPath } from './engine/utils'
export {
  toTextRunProperties, toDocxAlignment, toDocxBorder, toDocxShading,
  defaultTitleStyle, defaultSectionTitleStyle,
  defaultHeaderCellStyle, defaultCellStyle, defaultLabelStyle,
} from './engine/styler'
export { buildCover } from './engine/cover'
export { buildParagraph } from './engine/paragraph'
export { renderChart } from './engine/chart-renderer'
export { buildHeader, buildFooter } from './engine/header-footer'
export { mergeTheme, defaultTheme } from './engine/theme'

// ── Composables ────────────────────────────────────────────────────
export { useWordExporter } from './composables/useWordExporter'

// ── Components ─────────────────────────────────────────────────────
export { default as WordExport } from './components/WordExport.vue'
