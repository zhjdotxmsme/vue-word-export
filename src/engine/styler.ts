import {
  AlignmentType,
  BorderStyle as DocxBorderStyle,
  ShadingType,
  UnderlineType,
} from 'docx'
import type { TextRun as TextRunProperties } from 'docx'
import type { TextStyle, CellStyle, BorderStyle } from '../types/style'

/**
 * 将十六进制颜色（如 '#FF0000'）去除 #，docx 使用 RRGGBB 格式
 */
function stripHash(hex?: string): string | undefined {
  if (!hex) return undefined
  return hex.replace('#', '').trim()
}

/**
 * 将我们的 TextStyle 映射为 docx TextRun 的属性
 */
export function toTextRunProperties(style?: Partial<TextStyle>): Partial<TextRunProperties> {
  if (!style) return {}

  const props: Record<string, unknown> = {}

  if (style.bold !== undefined) props.bold = style.bold
  if (style.italic !== undefined) props.italic = style.italic
  if (style.fontSize !== undefined) props.size = style.fontSize * 2 // pt → half-points
  if (style.fontName !== undefined) props.font = { name: style.fontName }
  if (style.fontColor !== undefined) props.color = stripHash(style.fontColor)
  if (style.underline !== undefined) {
    props.underline = style.underline ? { type: UnderlineType.SINGLE } : undefined
  }

  return props as Partial<TextRunProperties>
}

/**
 * 将我们的 Alignment 映射为 docx AlignmentType
 */
export function toDocxAlignment(
  align?: 'left' | 'center' | 'right',
) {
  if (!align) return undefined
  if (align === 'left') return AlignmentType.LEFT
  if (align === 'center') return AlignmentType.CENTER
  return AlignmentType.RIGHT
}

/**
 * 将我们的 BorderStyle 映射为 docx 边框对象
 */
interface DocxBorderResult {
  style: string
  color?: string
  size?: number
}

export function toDocxBorder(
  border?: boolean | BorderStyle,
): DocxBorderResult | undefined {
  if (border === undefined || border === false) return undefined
  if (border === true) {
    return { style: DocxBorderStyle.SINGLE, size: 4 }
  }
  const styleMap: Record<string, string> = {
    single: DocxBorderStyle.SINGLE,
    double: DocxBorderStyle.DOUBLE,
    dotted: DocxBorderStyle.DOTTED,
  }
  return {
    style: styleMap[border.style] || DocxBorderStyle.SINGLE,
    color: border.color ? stripHash(border.color) : undefined,
    size: 4,
  }
}

/**
 * 创建单元格背景色着色对象
 */
export function toDocxShading(bgColor?: string): { type: string; fill: string } | undefined {
  if (!bgColor) return undefined
  return { type: ShadingType.CLEAR, fill: stripHash(bgColor) || '' }
}

/** 默认标题样式 */
export function defaultTitleStyle(): Partial<TextStyle> {
  return { bold: true, fontSize: 22, fontName: '微软雅黑' }
}

/** 默认区域小标题样式 */
export function defaultSectionTitleStyle(): Partial<TextStyle> {
  return { bold: true, fontSize: 14, fontName: '微软雅黑' }
}

/** 默认表头样式 */
export function defaultHeaderCellStyle(): Partial<CellStyle> {
  return {
    bold: true,
    fontSize: 10,
    fontName: '微软雅黑',
    bgColor: '#D9E2F3',
    align: 'center',
    border: true,
  }
}

/** 默认数据格样式 */
export function defaultCellStyle(): Partial<CellStyle> {
  return {
    fontSize: 10,
    fontName: '微软雅黑',
    border: true,
  }
}

/** 默认标签样式（基本信息表左侧） */
export function defaultLabelStyle(): Partial<CellStyle> {
  return {
    bold: true,
    fontSize: 10,
    fontName: '微软雅黑',
    bgColor: '#F2F2F2',
    border: true,
  }
}
