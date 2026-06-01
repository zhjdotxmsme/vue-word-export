import type { DocumentTheme, TextStyle, ParagraphStyle, CellStyle, ChartStyle, CoverStyle } from '../types/style'

/**
 * 内置默认主题
 */
export function defaultTheme(): Required<DocumentTheme> {
  return {
    fontFamily: '微软雅黑',
    fontSize: 10.5,
    title: {
      style: { bold: true, fontSize: 22, fontName: '微软雅黑' },
    },
    cover: {
      backgroundColor: '#FFFFFF',
      verticalAlign: 'middle',
      itemSpacing: 200,
      margins: undefined as any,
    },
    paragraph: {
      style: {
        align: 'left' as const,
        spacingBefore: undefined as any,
        spacingAfter: 100,
        lineSpacing: 360,
        indentFirstLine: 480,
        indentLeft: undefined as any,
        indentRight: undefined as any,
        headingLevel: undefined as any,
        pageBreakBefore: undefined as any,
        keepWithNext: undefined as any,
        keepLines: undefined as any,
      },
      heading: {
        1: { bold: true, fontSize: 18, fontName: '微软雅黑' },
        2: { bold: true, fontSize: 16, fontName: '微软雅黑' },
        3: { bold: true, fontSize: 14, fontName: '微软雅黑' },
      },
    },
    table: {
      basic: {
        labelStyle: { bold: true, fontSize: 10, fontName: '微软雅黑', bgColor: '#F2F2F2' } as any,
        valueStyle: { fontSize: 10, fontName: '微软雅黑' } as any,
        border: false,
      },
      list: {
        headerStyle: { bold: true, fontSize: 10, fontName: '微软雅黑', bgColor: '#D9E2F3', align: 'center' as const, border: true } as any,
        cellStyle: { fontSize: 10, fontName: '微软雅黑', border: true } as any,
        border: true,
      },
    },
    chart: {
      colorPalette: ['#1565C0', '#7B1FA2', '#00897B', '#E65100', '#C62828', '#F9A825'],
      style: {
        showLegend: true,
        legendPosition: 'bottom' as const,
        showDataLabel: false,
        colorPalette: undefined as any,
        xAxisName: undefined as any,
        yAxisName: undefined as any,
        backgroundColor: undefined as any,
      },
    },
    headerFooter: {
      style: { fontSize: 9, fontName: '微软雅黑', fontColor: '#999999' },
      border: true,
    },
  }
}

/**
 * 合并用户主题与默认主题（用户覆盖优先，深层合并）
 */
export function mergeTheme(userTheme?: DocumentTheme): Required<DocumentTheme> {
  const base = defaultTheme()
  if (!userTheme) return base

  function deepMerge<T>(target: T, source: Partial<T>): T {
    if (!source) return target
    const result = { ...target } as Record<string, unknown>
    for (const key of Object.keys(source) as (keyof T)[]) {
      const val = source[key]
      if (val !== undefined) {
        const targetVal = target[key]
        if (
          typeof val === 'object' && val !== null && !Array.isArray(val) &&
          typeof targetVal === 'object' && targetVal !== null && !Array.isArray(targetVal)
        ) {
          result[key as string] = deepMerge(targetVal as Record<string, unknown>, val as Record<string, unknown>)
        } else {
          result[key as string] = val as unknown
        }
      }
    }
    return result as unknown as T
  }

  return deepMerge(base, userTheme as unknown as Required<DocumentTheme>)
}
