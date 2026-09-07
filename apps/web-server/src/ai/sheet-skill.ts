/**
 * Sheet-AI skill handlers — formula suggest/explain, data analyze,
 * trend predict, chart suggest/create, data clean/fill. Placeholder
 * implementations carried over from the legacy `index.ts`.
 */
import { registerHandle } from '../common/index.js'

export function registerSheetAiSkillHandlers(): void {
  registerHandle('ai:sheets-formula-suggest', async (_event: unknown, args: unknown) => {
    const { dataRange, intent, sampleData } = args as {
      dataRange: string
      intent?: string
      sampleData?: unknown[][]
    }

    return {
      formulas: [
        {
          formula: '=SUM(A1:A10)',
          description: '求和公式，计算 A1 到 A10 的总和',
          score: 0.95,
        },
        {
          formula: '=AVERAGE(A1:A10)',
          description: '平均值公式，计算 A1 到 A10 的平均值',
          score: 0.90,
        },
        {
          formula: '=IF(A1>100,"高","低")',
          description: '条件判断，根据值返回不同结果',
          score: 0.85,
        },
      ],
      recommended: 0,
      range: dataRange,
    }
  })

  registerHandle('ai:sheets-formula-explain', async (_event: unknown, args: unknown) => {
    const { formula } = args as { formula: string }

    const explanations: Record<string, string> = {
      SUM: '对指定范围的所有数值求和',
      AVERAGE: '计算指定范围的算术平均值',
      IF: '根据条件返回不同的值',
      VLOOKUP: '在表格中垂直查找并返回对应值',
      INDEX: '返回指定行列交叉处的值',
      MATCH: '返回指定值在范围中的位置',
    }

    const funcName = formula.match(/=([A-Z]+)/)?.[1] || ''

    return {
      formula,
      function: funcName,
      explanation: explanations[funcName] || '未知公式',
      example: formula.replace(/[A-Z]+:/, 'A1:A10'),
      parameters: ['参数1', '参数2'],
    }
  })

  registerHandle('ai:sheets-data-analyze', async (_event: unknown, args: unknown) => {
    const { dataRange, sampleData } = args as {
      dataRange: string
      sampleData?: unknown[][]
    }

    return {
      summary: {
        rowCount: 100,
        columnCount: 5,
        numericColumns: ['A', 'C', 'E'],
        textColumns: ['B', 'D'],
      },
      insights: [
        { type: 'trend', description: 'C列呈上升趋势', confidence: 0.85 },
        { type: 'outlier', description: 'A列发现3个异常值', confidence: 0.78 },
        { type: 'correlation', description: 'A列与C列正相关', confidence: 0.82 },
      ],
      recommendations: [
        '建议添加趋势线',
        '考虑过滤异常值',
        '可以创建数据透视表',
      ],
    }
  })

  registerHandle('ai:sheets-trend-predict', async (_event: unknown, args: unknown) => {
    const { dataRange, periods } = args as {
      dataRange: string
      periods?: number
    }

    const predictPeriods = periods || 3
    return {
      predictions: Array.from({ length: predictPeriods }, (_, i) => ({
        period: i + 1,
        value: 100 + Math.random() * 20,
        lower: 90 + Math.random() * 10,
        upper: 110 + Math.random() * 10,
      })),
      model: 'linear_regression',
      confidence: 0.82,
    }
  })

  registerHandle('ai:sheets-chart-suggest', async (_event: unknown, args: unknown) => {
    const { dataRange, dataType } = args as {
      dataRange: string
      dataType?: 'categorical' | 'time_series' | 'numeric'
    }

    const chartTypes = {
      categorical: ['bar', 'column', 'pie'],
      time_series: ['line', 'area', 'combo'],
      numeric: ['scatter', 'bubble', 'histogram'],
    }

    const suggested = dataType ? chartTypes[dataType] : chartTypes.numeric

    return {
      suggestions: suggested.map((type, i) => ({
        type,
        score: 1 - i * 0.15,
        reason: `${type}图表最适合展示此类数据`,
      })),
      recommended: suggested[0],
    }
  })

  registerHandle('ai:sheets-chart-create', async (_event: unknown, args: unknown) => {
    const { dataRange, chartType, options } = args as {
      dataRange: string
      chartType: string
      options?: { title?: string; colors?: string[] }
    }

    return {
      chartId: `chart-${Date.now()}`,
      type: chartType,
      title: options?.title || 'AI 生成的图表',
      dataRange,
      options: {
        showLegend: true,
        showGrid: true,
        colors: options?.colors || ['#3498db', '#e74c3c', '#2ecc71'],
      },
    }
  })

  registerHandle('ai:sheets-data-clean', async (_event: unknown, args: unknown) => {
    const { dataRange } = args as { dataRange: string }

    return {
      issues: [
        { type: 'empty_cells', count: 5, suggestion: '填充或删除空单元格' },
        { type: 'duplicates', count: 3, suggestion: '删除重复行' },
        { type: 'formatting', count: 2, suggestion: '统一数字格式' },
      ],
      actions: [
        { operation: 'fill_empty', range: 'A1:A10', value: 'N/A' },
        { operation: 'remove_duplicates', range: 'A1:E100' },
        { operation: 'format_numbers', range: 'C1:C100', format: '#,##0.00' },
      ],
    }
  })

  registerHandle('ai:sheets-data-fill', async (_event: unknown, args: unknown) => {
    const { range, pattern } = args as {
      range: string
      pattern?: 'sequence' | 'copy' | 'formula'
    }

    return {
      filled: true,
      range,
      pattern: pattern || 'copy',
      values: ['值1', '值2', '值3', '值4', '值5'],
    }
  })
}
