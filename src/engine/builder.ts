import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  AlignmentType,
  PageOrientation,
  Header,
  Footer,
} from 'docx'
import type { ExportOptions } from '../types/export'
import { buildSection } from './table'
import { toTextRunProperties } from './styler'
import { downloadBlob } from './utils'
import { buildCover } from './cover'
import { buildParagraph } from './paragraph'
import { renderChart } from './chart-renderer'
import { buildHeader, buildFooter } from './header-footer'
import { mergeTheme } from './theme'

/**
 * 构建 Word 文档并返回 Blob
 */
export async function buildDocument(options: ExportOptions): Promise<Blob> {
  const { title, sections, data, pageOrientation, header, footer, differentFirstPage, theme: userTheme } = options

  const theme = mergeTheme(userTheme)
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
            size: 44,
            font: { name: '微软雅黑' },
            ...(titleStyle as Record<string, unknown>),
          }),
        ],
      }),
    )
  }

  // 2. 遍历 sections 并按 type 路由到对应渲染器
  let firstSection = true
  for (let i = 0; i < sections.length; i++) {
    const section = sections[i]
    let sectionChildren: (Paragraph | import('docx').Table)[] = []

    switch (section.type) {
      case 'cover':
        sectionChildren = buildCover(section, theme)
        break
      case 'paragraph':
        sectionChildren = buildParagraph(section, theme)
        break
      case 'chart': {
        const chartResult = await renderChart(section, data, theme.chart.colorPalette)
        sectionChildren = []
        if (section.title) {
          sectionChildren.push(
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { after: 100 },
              children: [new TextRun({ text: section.title, bold: true, size: 26 })],
            }),
          )
        }
        // 图表占位文本（真实 ImageRun 需异步加载图片二进制）
        sectionChildren.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 100, after: 100 },
            children: [
              new TextRun({
                text: `[${chartResult.chartType.toUpperCase()} 图表]`,
                italics: true,
                color: '666666',
              }),
            ],
          }),
        )
        sectionChildren.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: `数据条数: ${chartResult.svg ? '已生成' : '—'}`,
                size: 18,
                color: '999999',
              }),
            ],
          }),
        )
        break
      }
      case 'basic':
      case 'list':
        sectionChildren = buildSection(section, data)
        break
    }

    // Section 之间加间距（第一个元素前不加）
    if (!firstSection && section.type !== 'cover') {
      children.push(new Paragraph({ spacing: { before: 100 }, children: [] }))
    }

    children.push(...sectionChildren)
    firstSection = false

    // 封面之后加分页符
    if (section.type === 'cover') {
      children.push(new Paragraph({ pageBreakBefore: true, children: [] }))
    }
  }

  // 3. 构建 Section 级页眉/页脚
  const sectionHeaders: { default?: Header; first?: Header } = {}
  const sectionFooters: { default?: Footer; first?: Footer } = {}

  if (header) {
    const h = buildHeader(header, theme.headerFooter.style)
    if (h) sectionHeaders.default = h
    if (differentFirstPage) {
      // 首页（封面）使用空页眉
      sectionHeaders.first = new Header({ children: [] })
    }
  }

  if (footer) {
    const f = buildFooter(footer, theme.headerFooter.style)
    if (f) sectionFooters.default = f
    if (differentFirstPage) {
      sectionFooters.first = new Footer({ children: [] })
    }
  }

  // 4. 构建 Document
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
          ...(differentFirstPage ? { titlePage: true } : {}),
        },
        ...(Object.keys(sectionHeaders).length > 0 ? { headers: sectionHeaders } : {}),
        ...(Object.keys(sectionFooters).length > 0 ? { footers: sectionFooters } : {}),
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
