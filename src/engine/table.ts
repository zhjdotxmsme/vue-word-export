import {
  Table,
  TableRow,
  TableCell,
  Paragraph,
  TextRun,
  AlignmentType,
  BorderStyle,
} from 'docx'
import type { FieldConfig, BasicSection, ListSection, SectionConfig } from '../types/document'
import type { CellStyle, TableStyle } from '../types/style'
import {
  toDocxAlignment,
  defaultHeaderCellStyle,
  defaultCellStyle,
  defaultLabelStyle,
} from './styler'
import { getValueByPath } from './utils'

// ── Style merging ─────────────────────────────────────────────────────

function mergeCellStyles(
  ...styles: (Partial<CellStyle> | undefined)[]
): Partial<CellStyle> {
  const result: Partial<CellStyle> = {}
  for (const s of styles) {
    if (!s) continue
    if (s.bold !== undefined) result.bold = s.bold
    if (s.fontSize !== undefined) result.fontSize = s.fontSize
    if (s.fontColor !== undefined) result.fontColor = s.fontColor
    if (s.fontName !== undefined) result.fontName = s.fontName
    if (s.bgColor !== undefined) result.bgColor = s.bgColor
    if (s.align !== undefined) result.align = s.align
    if (s.verticalAlign !== undefined) result.verticalAlign = s.verticalAlign
    if (s.border !== undefined) result.border = s.border
  }
  return result
}

// ── Cell helpers ──────────────────────────────────────────────────────

function textRunOpts(style: Partial<CellStyle>): Record<string, unknown> {
  const opts: Record<string, unknown> = {}
  if (style.bold) opts.bold = true
  if (style.fontSize) opts.size = style.fontSize * 2
  if (style.fontColor) opts.color = style.fontColor.replace('#', '')
  if (style.fontName) opts.font = { name: style.fontName }
  return opts
}

function createParagraph(text: string, cellStyle: Partial<CellStyle>): Paragraph {
  return new Paragraph({
    alignment: toDocxAlignment(cellStyle.align) || AlignmentType.LEFT,
    spacing: { before: 30, after: 30 },
    children: [new TextRun({ text, ...textRunOpts(cellStyle) })],
  })
}

function createCell(
  children: Paragraph[],
  opts?: {
    width?: number | string
    borders?: boolean
    bgColor?: string
    verticalAlign?: 'top' | 'middle' | 'bottom'
  },
): TableCell {
  const cellBorders = opts?.borders
    ? {
        top: { style: BorderStyle.SINGLE, size: 4 },
        bottom: { style: BorderStyle.SINGLE, size: 4 },
        left: { style: BorderStyle.SINGLE, size: 4 },
        right: { style: BorderStyle.SINGLE, size: 4 },
      }
    : undefined

  const shading = opts?.bgColor
    ? { type: 'clear' as const, fill: opts.bgColor.replace('#', '') }
    : undefined

  const width = opts?.width !== undefined
    ? { size: typeof opts.width === 'number' ? opts.width : 3000, type: 'dxa' as const }
    : undefined

  return new TableCell({
    children,
    ...(cellBorders ? { borders: cellBorders } : {}),
    ...(shading ? { shading } : {}),
    ...(width ? { width } : {}),
  })
}

/**
 * 构建单个字段内容：label（加粗）+ value（常规）
 */
function buildFieldParagraph(
  field: FieldConfig,
  data: Record<string, unknown>,
  style: Partial<CellStyle>,
): Paragraph {
  const rawValue = getValueByPath(data, field.field)
  const text = field.format
    ? field.format(rawValue, data)
    : String(rawValue ?? '')

  return new Paragraph({
    spacing: { before: 30, after: 30 },
    children: [
      new TextRun({ text: field.label + '：', ...textRunOpts(style) }),
      new TextRun({ text, ...textRunOpts(mergeCellStyles(style, { bold: false })) }),
    ],
  })
}

// ═════════════════════════════════════════════════════════════════════
//  Basic Section — 多列并排（无边框表格）
// ═════════════════════════════════════════════════════════════════════

/**
 * 构建基本信息区（无边框表格，视觉上如同段落并排）
 *
 * - 每行 columns 个字段（默认 2），超出自动换行保持列内对齐
 * - 无边框、无背景色，视觉上与段落排版无异
 */
export function buildBasicTable(
  section: BasicSection,
  data: Record<string, unknown>,
): Table {
  const { fields, columns = 2 } = section
  const defaultStyle = defaultLabelStyle()

  // 每列等宽
  const colPercent = Math.floor(100 / columns)

  const rows: TableRow[] = []

  for (let i = 0; i < fields.length; i += columns) {
    const cells: TableCell[] = []

    for (let j = 0; j < columns; j++) {
      const field = fields[i + j]
      if (field) {
        const fieldStyle = mergeCellStyles(defaultStyle, field.style)
        const paragraph = buildFieldParagraph(field, data, fieldStyle)
        cells.push(
          createCell([paragraph], {
            width: `${colPercent}%`,
            verticalAlign: 'top',
          }),
        )
      } else {
        // 补齐空单元格
        cells.push(createCell([new Paragraph({ children: [] })], {
          width: `${colPercent}%`,
        }))
      }
    }

    rows.push(new TableRow({ children: cells }))
  }

  return new Table({
    rows,
    width: { size: 100, type: 'pct' },
    borders: {
      insideHorizontal: { style: BorderStyle.NONE, size: 0 },
      insideVertical: { style: BorderStyle.NONE, size: 0 },
      top: { style: BorderStyle.NONE, size: 0 },
      bottom: { style: BorderStyle.NONE, size: 0 },
      left: { style: BorderStyle.NONE, size: 0 },
      right: { style: BorderStyle.NONE, size: 0 },
    },
  })
}

// ═════════════════════════════════════════════════════════════════════
//  List Section — 数据列表（有边框表格）
// ═════════════════════════════════════════════════════════════════════

/**
 * 构建列表数据表格（表头 + 数据行，有边框）
 */
export function buildListTable(
  section: ListSection,
  data: Record<string, unknown>,
): Table {
  const { columns, dataField } = section
  const sectionStyle = section.style || {}
  const defaultHeader = mergeCellStyles(defaultHeaderCellStyle(), sectionStyle.headerStyle)
  const defaultDataStyle = mergeCellStyles(defaultCellStyle(), sectionStyle.cellStyle)

  const rows: TableRow[] = []

  // 1. 表头行
  const headerCells = columns.map((col) => {
    const style = mergeCellStyles(defaultHeader, col.style)
    return createCell([createParagraph(col.label, style)], {
      borders: true,
      bgColor: style.bgColor,
      width: col.width,
    })
  })
  rows.push(new TableRow({ children: headerCells }))

  // 2. 数据行
  const items = getValueByPath(data, dataField)
  if (Array.isArray(items)) {
    for (const item of items) {
      const rowCells = columns.map((col) => {
        const rawValue = getValueByPath(item as Record<string, unknown>, col.field)
        const text = col.format
          ? col.format(rawValue, item as Record<string, unknown>)
          : String(rawValue ?? '')
        const style = mergeCellStyles(defaultDataStyle, col.style, { align: col.align })
        return createCell([createParagraph(text, style)], {
          borders: true,
          width: col.width,
        })
      })
      rows.push(new TableRow({ children: rowCells }))
    }
  }

  return new Table({
    rows,
    width: { size: 100, type: 'pct' },
  })
}

// ═════════════════════════════════════════════════════════════════════
//  Section builder
// ═════════════════════════════════════════════════════════════════════

/**
 * 构建 Section 段落（标题 + 内容）
 *
 * - basic: 无边框多列表格
 * - list: 有边框数据表格
 */
export function buildSection(
  section: SectionConfig,
  data: Record<string, unknown>,
): (Paragraph | Table)[] {
  const children: (Paragraph | Table)[] = []

  if (section.title) {
    children.push(
      new Paragraph({
        spacing: { before: 200, after: 100 },
        children: [
          new TextRun({
            text: section.title,
            bold: true,
            size: 28,
            font: { name: '微软雅黑' },
          }),
        ],
      }),
    )
  }

  if (section.type === 'basic') {
    children.push(buildBasicTable(section, data))
  } else if (section.type === 'list') {
    children.push(buildListTable(section, data))
  }

  return children
}
