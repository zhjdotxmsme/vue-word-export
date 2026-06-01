import type { CellStyle, TableStyle } from './style'

/** 字段配置——描述单个字段在文档中的映射 */
export interface FieldConfig {
  /** 显示标签（如 "企业名称"） */
  label: string
  /** 数据字段路径，支持点号嵌套（如 "company.name"） */
  field: string
  /** 列宽（磅值如 3000，或百分比字符串如 "20%"） */
  width?: number | string
  /** 对齐方式 */
  align?: 'left' | 'center' | 'right'
  /** 单元格样式 */
  style?: Partial<CellStyle>
  /** 值格式化函数 */
  format?: (value: unknown, row: Record<string, unknown>) => string
}

/** 基本信息区——键值对表格 */
export interface BasicSection {
  type: 'basic'
  /** 区域标题（可选小标题） */
  title?: string
  /** 按序显示的字段 */
  fields: FieldConfig[]
  /** 表格列数：1=单列 (label: value)，2=双列 (label | value + label | value)，默认 2 */
  columns?: 1 | 2
  /** 表格样式 */
  style?: Partial<TableStyle>
}

/** 列表区——数据数组表格 */
export interface ListSection {
  type: 'list'
  /** 列表标题（可选） */
  title?: string
  /** 数组数据字段路径 */
  dataField: string
  /** 列配置 */
  columns: FieldConfig[]
  /** 表格样式 */
  style?: Partial<TableStyle>
}

export type SectionConfig = BasicSection | ListSection
