import { Paragraph, TextRun, AlignmentType } from 'docx'
import type { CoverSection, CoverItem, CoverTitle, CoverSubtitle, CoverDate, CoverText } from '../types/document'
import type { CoverStyle } from '../types/style'
import { toTextRunProperties } from './styler'

/**
 * 构建封面段落列表
 *
 * @param section 封面配置段
 * @param theme   已合并的主题（只需 cover / title）
 * @returns 封面段落数组
 */
export function buildCover(
  section: CoverSection,
  theme: {
    cover?: Partial<CoverStyle>
    title?: { style?: Record<string, unknown> }
  },
): Paragraph[] {
  const children: Paragraph[] = []

  // 合并 section.style 与 theme.cover，section.style 优先
  const mergedStyle: CoverStyle = {
    itemSpacing: 200,
    verticalAlign: 'middle',
    ...theme.cover,
    ...section.style,
  }

  const { itemSpacing = 200, verticalAlign } = mergedStyle

  // 垂直居中：在内容前插入一个 spacer 段落
  if (verticalAlign === 'middle') {
    children.push(new Paragraph({ spacing: { before: 3000 }, children: [] }))
  }

  for (const item of section.items) {
    switch (item.type) {
      case 'title':
        children.push(buildTitleParagraph(item, theme))
        break
      case 'subtitle':
        children.push(buildSubtitleParagraph(item))
        break
      case 'date':
        children.push(buildDateParagraph(item))
        break
      case 'text':
        children.push(buildTextParagraph(item))
        break
      case 'image':
        children.push(buildImagePlaceholder())
        break
    }

    // 项间间距
    children.push(new Paragraph({ spacing: { before: itemSpacing }, children: [] }))
  }

  return children
}

// ── 内部构建函数 ──────────────────────────────────────────────────────

/**
 * 构建封面标题段落
 */
function buildTitleParagraph(
  item: CoverTitle,
  theme: { title?: { style?: Record<string, unknown> } },
): Paragraph {
  const text = typeof item.text === 'function' ? item.text() : item.text
  const mergedStyle = { ...theme.title?.style, ...item.style }
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 400, after: 200 },
    children: [
      new TextRun({
        text,
        bold: true,
        size: 44,
        ...toTextRunProperties(mergedStyle),
      }),
    ],
  })
}

/**
 * 构建封面副标题段落
 */
function buildSubtitleParagraph(item: CoverSubtitle): Paragraph {
  const text = typeof item.text === 'function' ? item.text() : item.text
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 100 },
    children: [new TextRun({ text, ...toTextRunProperties(item.style) })],
  })
}

/**
 * 构建封面日期段落
 */
function buildDateParagraph(item: CoverDate): Paragraph {
  const fmt = item.format || 'YYYY-MM-DD'
  const now = new Date()
  const dateStr = fmt
    .replace('YYYY', String(now.getFullYear()))
    .replace('MM', String(now.getMonth() + 1).padStart(2, '0'))
    .replace('DD', String(now.getDate()).padStart(2, '0'))

  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200 },
    children: [new TextRun({ text: dateStr, ...toTextRunProperties(item.style) })],
  })
}

/**
 * 构建封面自定义文本段落
 */
function buildTextParagraph(item: CoverText): Paragraph {
  const text = typeof item.text === 'function' ? item.text() : item.text
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 100 },
    children: [new TextRun({ text, ...toTextRunProperties(item.style) })],
  })
}

/**
 * 构建封面图片占位段落
 *
 * 真实图片需要异步加载 base64/URL 数据，
 * 此处先用占位文本标记插入位置，由上层决定是否替换为 ImageRun。
 */
function buildImagePlaceholder(): Paragraph {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 200 },
    children: [new TextRun({ text: '[图片占位]', italics: true, color: '999999' })],
  })
}
