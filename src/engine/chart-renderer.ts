import type { ChartSection } from '../types/document'
import type { ChartStyle, ChartDimension, ChartValueField } from '../types/style'

export interface ChartRenderResult {
  base64: string
  svg: string
  width: number
  height: number
  chartType: string
}

/**
 * 渲染图表为 base64 SVG
 *
 * 使用 ECharts SVG 渲染器生成矢量图，无需 Canvas/Node-Canvas 依赖。
 * ECharts 通过 dynamic import 延迟加载，仅在首次使用 ChartSection 时加载。
 */
export async function renderChart(
  section: ChartSection,
  data: Record<string, unknown>,
  colorPalette?: string[],
): Promise<ChartRenderResult> {
  const items = getValueByPath(data, section.dataField)
  if (!Array.isArray(items) || items.length === 0) {
    throw new Error(`Chart data at '${section.dataField}' 为空或不是数组`)
  }

  const width = section.width || 8000
  const height = section.height || 3000

  // Build ECharts option from configuration
  const option = buildChartOption(section, items as Record<string, unknown>[], colorPalette)

  // Dynamic import of echarts (optional peer dependency)
  let echarts: any
  try {
    echarts = await import('echarts')
  } catch {
    throw new Error(
      'echarts 是 ChartSection 的可选依赖，请安装: npm install echarts',
    )
  }

  // Use SVG renderer (no DOM needed)
  const chart = echarts.init(null, null, {
    renderer: 'svg',
    width: Math.round(width / 1440 * 96),
    height: Math.round(height / 1440 * 96),
  })

  chart.setOption(option)
  const svgStr: string = chart.renderToSVGString()
  chart.dispose()

  return {
    svg: svgStr,
    base64: Buffer.from(svgStr).toString('base64'),
    width,
    height,
    chartType: section.chartType,
  }
}

/**
 * 从数据和配置构建 ECharts option 对象
 */
function buildChartOption(
  section: ChartSection,
  items: Record<string, unknown>[],
  palette?: string[],
): Record<string, unknown> {
  const { chartType, dimension, style } = section
  const colors = style?.colorPalette || palette || [
    '#1565C0', '#7B1FA2', '#00897B', '#E65100', '#C62828', '#F9A825',
  ]

  const categories = items.map((item) => String(item[dimension.categoryField]))

  if (chartType === 'pie') {
    const seriesData = items.map((item, i) => ({
      name: String(item[dimension.categoryField]),
      value: Number(item[dimension.valueFields[0].field]),
      itemStyle: { color: colors[i % colors.length] },
    }))

    return {
      tooltip: { trigger: 'item', formatter: '{b}: {c} ({d}%)' },
      legend: style?.showLegend !== false
        ? { bottom: style?.legendPosition === 'bottom' ? 0 : undefined, left: 'center' }
        : undefined,
      series: [{
        type: 'pie',
        radius: ['30%', '60%'],
        center: ['50%', '50%'],
        avoidLabelOverlap: true,
        label: { show: style?.showDataLabel ?? false, formatter: '{b}: {d}%' },
        emphasis: { label: { show: true, fontSize: 14, fontWeight: 'bold' } },
        data: seriesData,
      }],
      color: colors,
    }
  }

  // Line / Bar / Area
  const series = dimension.valueFields.map((vf, i) => {
    const dataValues = items.map((item) => Number(item[vf.field]))
    const s: Record<string, unknown> = {
      name: vf.name || vf.field,
      type: chartType === 'area' ? 'line' : chartType,
      data: dataValues,
      itemStyle: { color: vf.color || colors[i % colors.length] },
      smooth: chartType === 'line',
    }
    if (chartType === 'area') {
      s.areaStyle = {}
    }
    return s
  })

  return {
    tooltip: { trigger: 'axis' },
    legend: style?.showLegend !== false
      ? { data: dimension.valueFields.map((v) => v.name || v.field), bottom: 0 }
      : undefined,
    grid: { left: '3%', right: '4%', bottom: style?.showLegend !== false ? '15%' : '10%', top: '10%', containLabel: true },
    xAxis: {
      type: 'category',
      name: style?.xAxisName,
      data: categories,
      axisLabel: { rotate: categories.length > 8 ? 45 : 0 },
    },
    yAxis: {
      type: 'value',
      name: style?.yAxisName,
    },
    series,
    color: colors,
  }
}

/**
 * 从嵌套对象中按点号路径取值（等同于 utils.getValueByPath，避免交叉引用）
 */
function getValueByPath(obj: Record<string, unknown>, path: string): unknown {
  return path.split('.').reduce((acc: unknown, key: string) => {
    if (acc && typeof acc === 'object') {
      return (acc as Record<string, unknown>)[key]
    }
    return undefined
  }, obj as unknown)
}
