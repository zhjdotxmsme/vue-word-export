import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  AlignmentType,
  PageOrientation,
} from 'docx'
import type { ExportOptions } from '../types/export'
import { buildSection } from './table'
import { toTextRunProperties } from './styler'
import { downloadBlob } from './utils'

/**
 * 构建 Word 文档并返回 Blob
 */
export async function buildDocument(options: ExportOptions): Promise<Blob> {
  const { title, sections, data, pageOrientation } = options

  const children: (Paragraph | import('docx').Table)[] = []

  // 1. 文档标题（支持字符串或配置对象）
  if (title) {
    const resolved = typeof title === 'string' ? { text: title } : title
    const titleText = typeof resolved.text === 'function' ? resolved.text() : resolved.text
    const titleStyle = resolved.style ? toTextRunProperties(resolved.style) : {}

    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 400 },
        children: [
          new TextRun({
            text: titleText,
            bold: true,
            size: 44, // 22pt
            font: { name: '微软雅黑' },
            ...(titleStyle as Record<string, unknown>),
          }),
        ],
      }),
    )
  }

  // 2. 遍历 sections
  for (let i = 0; i < sections.length; i++) {
    const section = sections[i]
    const sectionChildren = buildSection(section, data)

    // Section 之间加间距
    if (i > 0) {
      children.push(
        new Paragraph({ spacing: { before: 100 }, children: [] }),
      )
    }

    children.push(...sectionChildren)
  }

  // 3. 构建 Document
  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            ...(pageOrientation === 'landscape'
              ? { orientation: PageOrientation.LANDSCAPE }
              : {}),
            margin: {
              top: options.margins?.top ?? 1440,
              bottom: options.margins?.bottom ?? 1440,
              left: options.margins?.left ?? 1800,
              right: options.margins?.right ?? 1800,
            },
          },
        },
        children,
      },
    ],
  })

  return await Packer.toBlob(doc)
}

/**
 * 导出 Word 文档（构建 + 下载一步完成）
 */
export async function exportWord(options: ExportOptions): Promise<void> {
  const blob = await buildDocument(options)
  const filename = options.filename || 'export.docx'
  downloadBlob(blob, filename)
}
