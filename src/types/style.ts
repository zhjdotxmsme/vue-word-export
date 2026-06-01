/** 文字样式（对应 docx 中 IRunProperties） */
export interface TextStyle {
  bold?: boolean
  italic?: boolean
  fontSize?: number
  fontColor?: string
  fontName?: string
  underline?: boolean
}

/** 单元格样式 */
export interface CellStyle {
  bold?: boolean
  fontSize?: number
  fontColor?: string
  fontName?: string
  bgColor?: string
  align?: 'left' | 'center' | 'right'
  verticalAlign?: 'top' | 'middle' | 'bottom'
  border?: boolean | BorderStyle
}

export interface BorderStyle {
  style: 'single' | 'double' | 'dotted'
  color?: string
}

/** 表格样式 */
export interface TableStyle {
  headerStyle?: Partial<CellStyle>
  cellStyle?: Partial<CellStyle>
  border?: boolean
  autoWidth?: boolean
}

// ── ParagraphStyle ────────────────────────────────────────────────────

export interface ParagraphStyle {
  align?: 'left' | 'center' | 'right' | 'justify'
  spacingBefore?: number
  spacingAfter?: number
  lineSpacing?: number
  indentFirstLine?: number
  indentLeft?: number
  indentRight?: number
  headingLevel?: 1 | 2 | 3 | 4 | 5 | 6
  pageBreakBefore?: boolean
  keepWithNext?: boolean
  keepLines?: boolean
}

// ── ChartStyle ────────────────────────────────────────────────────────

export interface ChartDimension {
  categoryField: string
  valueFields: ChartValueField[]
}

export interface ChartValueField {
  field: string
  name?: string
  color?: string
  format?: (value: unknown) => string
}

export interface ChartStyle {
  colorPalette?: string[]
  showLegend?: boolean
  legendPosition?: 'top' | 'bottom' | 'left' | 'right'
  xAxisName?: string
  yAxisName?: string
  showDataLabel?: boolean
  backgroundColor?: string
}

// ── CoverStyle ────────────────────────────────────────────────────────

export interface CoverStyle {
  backgroundColor?: string
  verticalAlign?: 'middle' | 'top' | 'bottom'
  itemSpacing?: number
  margins?: { top?: number; bottom?: number; left?: number; right?: number }
}

// ── DocumentTheme ─────────────────────────────────────────────────────

export interface DocumentTheme {
  fontFamily?: string
  fontSize?: number
  title?: { style?: Partial<TextStyle> }
  cover?: Partial<CoverStyle>
  paragraph?: {
    style?: Partial<ParagraphStyle>
    heading?: Record<number, Partial<TextStyle>>
  }
  table?: {
    basic?: { labelStyle?: Partial<CellStyle>; valueStyle?: Partial<CellStyle>; border?: boolean }
    list?: { headerStyle?: Partial<CellStyle>; cellStyle?: Partial<CellStyle>; border?: boolean }
  }
  chart?: { colorPalette?: string[]; style?: Partial<ChartStyle> }
  headerFooter?: { style?: Partial<TextStyle>; border?: boolean }
}
