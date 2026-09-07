/**
 * Mobile / multimodal — device detection and image-understanding
 * placeholders. Both intentionally return canned responses for now.
 */
import { registerHandle } from '../common/index.js'

export function registerMobileHandlers(): void {
  registerHandle('mobile:get-settings', () => ({
    touchEnabled: true,
    viewportWidth: 375,
    viewportHeight: 667,
    pixelRatio: 2,
    supportsTouch: true,
    supportsPen: true,
    deviceType: 'auto',
    theme: 'light',
    reducedMotion: false,
    darkMode: false,
  }))

  registerHandle('mobile:detect', () => ({
    isMobile: true,
    isTablet: false,
    isDesktop: false,
    os: 'web',
    browser: 'web',
    screenWidth: 375,
    screenHeight: 667,
    orientation: 'portrait',
  }))
}

export function registerMultimodalHandlers(): void {
  registerHandle('multimodal:analyze-image', async (_event: unknown, _args: unknown) => ({
    description: '图片内容分析：这是一张包含文字和图表的图片。',
    tags: ['图表', '文字', '数据'],
    text: '图片中包含数据可视化内容',
    confidence: 0.92,
    objects: [
      { label: '柱状图', confidence: 0.95, boundingBox: { x: 10, y: 10, width: 100, height: 100 } },
      { label: '标题', confidence: 0.88, boundingBox: { x: 10, y: 5, width: 80, height: 20 } },
    ],
  }))

  registerHandle('multimodal:extract-table', async (_event: unknown, _args: unknown) => ({
    rows: 5,
    columns: 4,
    headers: ['姓名', '年龄', '职位', '部门'],
    data: [
      ['张三', '28', '工程师', '研发部'],
      ['李四', '32', '经理', '产品部'],
    ],
  }))
}
