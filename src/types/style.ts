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
