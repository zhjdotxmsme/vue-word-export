import type { SectionConfig } from './document'
import type { TextStyle, DocumentTheme } from './style'

// ── Header / Footer ─────────────────────────────────────────────────

export interface HeaderFooterText {
  text: string | (() => string)
  style?: Partial<TextStyle>
}

export type HeaderFooterElement =
  | HeaderFooterText
  | { type: 'pageNumber' }
  | { type: 'totalPages' }
  | { type: 'date'; format?: string }
  | { type: 'image'; src: string; width?: number; height?: number }

export type HeaderFooterContent =
  | string
  | HeaderFooterText
  | HeaderFooterElement[]

export interface HeaderFooterConfig {
  left?: HeaderFooterContent
  center?: HeaderFooterContent
  right?: HeaderFooterContent
  border?: boolean
}

// ── ExportOptions ────────────────────────────────────────────────────

/** 导出选项 */
export interface ExportOptions {
  /** 文档标题（可传入字符串或配置对象） */
  title?: string | {
    text: string | (() => string)
    style?: Partial<TextStyle>
  }
  /** 文档正文段落（按序渲染） */
  sections: SectionConfig[]
  /** 后端 JSON 数据 */
  data: Record<string, unknown>
  /** 导出文件名（默认 'export.docx'） */
  filename?: string
  /** 页面方向 */
  pageOrientation?: 'portrait' | 'landscape'
  /** 页边距（twips，1 twip = 1/1440 inch） */
  margins?: {
    top?: number
    bottom?: number
    left?: number
    right?: number
  }
  /** 页眉配置 */
  header?: HeaderFooterConfig
  /** 页脚配置 */
  footer?: HeaderFooterConfig
  /** 首页不同（封面页自动不显示页眉页脚） */
  differentFirstPage?: boolean
  /** 全局样式主题 */
  theme?: DocumentTheme
}
