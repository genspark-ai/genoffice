/**
 * Chart generation + visualization dashboard. Generates SVG strings for
 * bar/pie charts (placeholder visualisations); dashboard endpoints return
 * skeleton structures.
 */
import { registerHandle } from '../common/index.js'

export function registerChartHandlers(): void {
  registerHandle('chart:generate', async (_event: unknown, args: unknown) => {
    const { type, data, options } = args as {
      type: 'bar' | 'line' | 'pie' | 'scatter' | 'radar'
      data: { labels?: string[]; datasets: Array<{ label: string; data: number[] }> }
      options?: { title?: string; colors?: string[] }
    }

    const colors = options?.colors || ['#4CAF50', '#2196F3', '#FF9800', '#E91E63', '#9C27B0']
    const width = 600
    const height = 400
    const padding = 50

    let svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}">`

    if (options?.title) {
      svg += `<text x="${width / 2}" y="30" text-anchor="middle" font-size="18" font-weight="bold">${options.title}</text>`
    }

    if (type === 'bar') {
      const labels = data.labels || []
      const barWidth = (width - padding * 2) / (labels.length || 1)
      const maxVal = Math.max(...(data.datasets[0]?.data || [100]))

      data.datasets.forEach((dataset, di) => {
        dataset.data.forEach((value, i) => {
          const x = padding + i * barWidth + barWidth * 0.1
          const barHeight = (value / maxVal) * (height - padding * 2)
          const y = height - padding - barHeight
          const color = colors[di % colors.length]

          svg += `<rect x="${x}" y="${y}" width="${barWidth * 0.8}" height="${barHeight}" fill="${color}" />`
          svg += `<text x="${x + barWidth * 0.4}" y="${height - padding + 20}" text-anchor="middle">${labels[i] || ''}</text>`
          svg += `<text x="${x + barWidth * 0.4}" y="${y - 5}" text-anchor="middle">${value}</text>`
        })
      })
    } else if (type === 'pie') {
      const total = (data.datasets[0]?.data || [1]).reduce((a, b) => a + b, 0)
      let currentAngle = 0

      ;(data.datasets[0]?.data || []).forEach((value, i) => {
        const angle = (value / total) * 360
        const startAngle = currentAngle
        const endAngle = currentAngle + angle
        currentAngle = endAngle

        const cx = width / 2
        const cy = height / 2
        const r = Math.min(width, height) / 2 - padding

        const x1 = cx + r * Math.cos((startAngle - 90) * Math.PI / 180)
        const y1 = cy + r * Math.sin((startAngle - 90) * Math.PI / 180)
        const x2 = cx + r * Math.cos((endAngle - 90) * Math.PI / 180)
        const y2 = cy + r * Math.sin((endAngle - 90) * Math.PI / 180)

        const largeArc = angle > 180 ? 1 : 0
        const color = colors[i % colors.length]

        svg += `<path d="M${cx},${cy} L${x1},${y1} A${r},${r} 0 ${largeArc},1 ${x2},${y2} Z" fill="${color}" />`
        svg += `<text x="${cx + r * 0.6 * Math.cos((startAngle + angle / 2 - 90) * Math.PI / 180)}" y="${cy + r * 0.6 * Math.sin((startAngle + angle / 2 - 90) * Math.PI / 180)}" text-anchor="middle" fill="white">${((value / total) * 100).toFixed(1)}%</text>`
      })
    }

    svg += '</svg>'

    return {
      svg,
      type,
      data,
      dimensions: { width, height },
    }
  })

  registerHandle('visualization:create-dashboard', async (_event: unknown, args: unknown) => {
    const { widgets } = args as {
      widgets: Array<{
        id: string
        type: 'chart' | 'table' | 'metric' | 'text'
        data: unknown
        position: { x: number; y: number; w: number; h: number }
      }>
    }

    return {
      id: `dashboard-${Date.now()}`,
      widgets: widgets.map(w => ({
        id: w.id,
        type: w.type,
        position: w.position,
      })),
      createdAt: Date.now(),
    }
  })

  registerHandle('visualization:get-chart-data', (_event: unknown, _args: unknown) => ({
    labels: ['一月', '二月', '三月', '四月', '五月'],
    datasets: [
      { label: '销售额', data: [120, 150, 180, 140, 200] },
      { label: '成本', data: [80, 90, 100, 85, 110] },
    ],
  }))
}
