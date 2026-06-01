import { Paragraph, TextRun, AlignmentType } from 'docx'
import type { ParagraphSection, TextRunFragment } from '../types/document'
import type { ParagraphStyle } from '../types/style'
import { toTextRunProperties } from './styler'

/**
 * 构建段落（支持纯文本和富文本混排）
 *
 * - content 为单元素时渲染为纯文本段落
 * - content 为数组时渲染为富文本段落（不同 TextRunFragment 不同样式）
 * - BreakFragment 触发段内换行（作为新段落处理）
 */
export function buildParagraph(
  section: ParagraphSection,
  theme?: { paragraph?: { style?: Partial<ParagraphStyle> } },
): (Paragraph | import('docx').Table)[] {
  const children: (Paragraph | import('docx').Table)[] = []

  // Section title（可选）
  if (section.title) {
    children.push(new Paragraph({
      spacing: { before: 200, after: 100 },
      children: [new TextRun({
        text: section.title,
        bold: true,
        size: 28,
        font: { name: '微软雅黑' },
      })],
    }))
  }

  // Merge paragraph styles: theme default → section style
  const mergedStyle: Partial<ParagraphStyle> = {
    ...(theme?.paragraph?.style || {}),
    ...section.style,
  }

  const contents = Array.isArray(section.content) ? section.content : [section.content]
  let runs: TextRun[] = []

  for (const fragment of contents) {
    if (fragment.type === 'break') {
      // Flush current paragraph
      if (runs.length > 0) {
        children.push(buildDocxParagraph(runs, mergedStyle))
        runs = []
      }
      // Add empty paragraph for line break
      children.push(buildDocxParagraph([], { ...mergedStyle, spacingBefore: 0, spacingAfter: 0 }))
      continue
    }

    const tf = fragment as TextRunFragment
    const text = typeof tf.text === 'function' ? tf.text() : tf.text
    const runProps = toTextRunProperties(tf.style)
    runs.push(new TextRun({ text, ...runProps }))
  }

  // Flush remaining runs
  if (runs.length > 0) {
    children.push(buildDocxParagraph(runs, mergedStyle))
  }

  return children
}

function buildDocxParagraph(runs: TextRun[], style: Partial<ParagraphStyle>): Paragraph {
  const alignmentMap: Record<string, typeof AlignmentType[keyof typeof AlignmentType]> = {
    left: AlignmentType.LEFT,
    center: AlignmentType.CENTER,
    right: AlignmentType.RIGHT,
    justify: AlignmentType.JUSTIFIED,
  }

  return new Paragraph({
    alignment: style.align ? alignmentMap[style.align] : AlignmentType.LEFT,
    spacing: {
      before: style.spacingBefore,
      after: style.spacingAfter,
      line: style.lineSpacing,
    },
    indent: {
      firstLine: style.indentFirstLine,
      left: style.indentLeft,
      right: style.indentRight,
    },
    children: runs,
  })
}
