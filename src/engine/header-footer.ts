import { Header, Footer, Paragraph, TextRun, AlignmentType } from 'docx'
import type { HeaderFooterConfig, HeaderFooterContent, HeaderFooterElement, HeaderFooterText } from '../types/export'
import type { TextStyle } from '../types/style'
import { toTextRunProperties } from './styler'

/**
 * 构建页眉
 */
export function buildHeader(
  config: HeaderFooterConfig,
  themeStyle?: Partial<TextStyle>,
): Header | undefined {
  if (!config) return undefined

  const paragraphs: Paragraph[] = []
  const defaultStyle = themeStyle || { fontSize: 9, fontName: '微软雅黑', fontColor: '#999999' }

  // Left content
  if (config.left) {
    paragraphs.push(new Paragraph({
      alignment: AlignmentType.LEFT,
      children: buildContent(config.left, defaultStyle),
    }))
  }

  // Center content
  if (config.center) {
    paragraphs.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      children: buildContent(config.center, defaultStyle),
    }))
  }

  // Right content
  if (config.right) {
    paragraphs.push(new Paragraph({
      alignment: AlignmentType.RIGHT,
      children: buildContent(config.right, defaultStyle),
    }))
  }

  return new Header({ children: paragraphs })
}

/**
 * 构建页脚
 */
export function buildFooter(
  config: HeaderFooterConfig,
  themeStyle?: Partial<TextStyle>,
): Footer | undefined {
  if (!config) return undefined

  const paragraphs: Paragraph[] = []
  const defaultStyle = themeStyle || { fontSize: 9, fontName: '微软雅黑', fontColor: '#999999' }

  if (config.left) {
    paragraphs.push(new Paragraph({
      alignment: AlignmentType.LEFT,
      children: buildContent(config.left, defaultStyle),
    }))
  }

  if (config.center) {
    paragraphs.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      children: buildContent(config.center, defaultStyle),
    }))
  }

  if (config.right) {
    paragraphs.push(new Paragraph({
      alignment: AlignmentType.RIGHT,
      children: buildContent(config.right, defaultStyle),
    }))
  }

  return new Footer({ children: paragraphs })
}

/**
 * 将 HeaderFooterContent 解析为 TextRun 数组
 */
function buildContent(
  content: HeaderFooterContent,
  defaultStyle: Partial<TextStyle>,
): TextRun[] {
  // Simple string
  if (typeof content === 'string') {
    return [new TextRun({ text: content, ...toTextRunProperties(defaultStyle) })]
  }

  // HeaderFooterText object (not array)
  if (!Array.isArray(content)) {
    const hft = content as HeaderFooterText
    const text = typeof hft.text === 'function' ? hft.text() : hft.text
    const mergedStyle = { ...defaultStyle, ...hft.style }
    return [new TextRun({ text, ...toTextRunProperties(mergedStyle) })]
  }

  // Array of HeaderFooterElements
  return content.map((element) => {
    if ('type' in element) {
      const el = element as any

      if (el.type === 'pageNumber') {
        // docx handles page numbers via fields, but for now use placeholder
        return new TextRun({
          text: '[页码]',
          ...toTextRunProperties({ ...defaultStyle }),
        })
      }

      if (el.type === 'totalPages') {
        return new TextRun({
          text: '[总页数]',
          ...toTextRunProperties({ ...defaultStyle }),
        })
      }

      if (el.type === 'date') {
        const fmt = el.format || 'YYYY-MM-DD'
        const now = new Date()
        const dateStr = fmt
          .replace('YYYY', String(now.getFullYear()))
          .replace('MM', String(now.getMonth() + 1).padStart(2, '0'))
          .replace('DD', String(now.getDate()).padStart(2, '0'))
        return new TextRun({
          text: dateStr,
          ...toTextRunProperties({ ...defaultStyle }),
        })
      }

      if (el.type === 'image') {
        return new TextRun({
          text: '[Logo]',
          ...toTextRunProperties({ ...defaultStyle, italic: true }),
        })
      }
    }

    // Regular HeaderFooterText in array
    const hft = element as HeaderFooterText
    const text = typeof hft.text === 'function' ? hft.text() : hft.text
    const mergedStyle = { ...defaultStyle, ...hft.style }
    return new TextRun({ text, ...toTextRunProperties(mergedStyle) })
  })
}
