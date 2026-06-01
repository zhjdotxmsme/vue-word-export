import type { SectionConfig } from './document'
import type { TextStyle } from './style'

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
}
