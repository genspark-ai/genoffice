/**
 * Offscreen frame source for the video export: one React root holding either the animated
 * page stage or the Morph tween, re-rendered synchronously per frame and read back as a canvas.
 */
import React from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { flushSync } from 'react-dom'
import type Konva from 'konva'
import type { RenderSlide } from '@genoffice/pptx-render'
import type { ShapeKey } from '../../shared/ipc'
import type { NodeAnimState } from '../animation-play'
import { AnimatedSlideStage } from '../components/AnimatedSlide'
import { MorphStage } from '../components/MorphStage'

export class FrameRenderer {
  private readonly container: HTMLDivElement
  private readonly root: Root
  private stage: Konva.Stage | null = null

  constructor(
    readonly width: number,
    readonly height: number,
    private readonly images: Map<string, HTMLImageElement>,
  ) {
    this.container = document.createElement('div')
    // display:none would give the Konva canvas zero size; park it off-viewport instead
    this.container.style.cssText = 'position:fixed;left:-100000px;top:0;pointer-events:none;'
    document.body.appendChild(this.container)
    this.root = createRoot(this.container)
  }

  /** The slide at one animation state, drawn into a fresh canvas of the output size. */
  slideFrame(slide: RenderSlide, states: Map<string, NodeAnimState>): HTMLCanvasElement {
    this.mount(
      <AnimatedSlideStage
        slide={slide}
        images={this.images}
        width={this.width}
        states={states}
        stageRef={(s) => (this.stage = s)}
      />,
    )
    return this.capture()
  }

  /** The Morph tween between two slides at raw progress 0..1. */
  morphFrame(
    from: RenderSlide,
    to: RenderSlide,
    fromKeys: ShapeKey[],
    toKeys: ShapeKey[],
    progress: number,
  ): HTMLCanvasElement {
    this.mount(
      <MorphStage
        from={from}
        to={to}
        fromKeys={fromKeys}
        toKeys={toKeys}
        images={this.images}
        width={this.width}
        progress={progress}
        onDone={() => undefined}
        stageRef={(s) => (this.stage = s)}
      />,
    )
    return this.capture()
  }

  private mount(el: React.ReactElement) {
    flushSync(() => this.root.render(el))
  }

  private capture(): HTMLCanvasElement {
    const stage = this.stage
    if (!stage) throw new Error('video export: stage did not mount')
    // react-konva schedules a batchDraw; the frame is needed now
    for (const layer of stage.getLayers()) layer.draw()
    // pixelRatio 1: the layer canvases are devicePixelRatio-sized, the output is not
    const drawn = stage.toCanvas({ pixelRatio: 1 })
    const out = document.createElement('canvas')
    out.width = this.width
    out.height = this.height
    const ctx = out.getContext('2d')
    if (!ctx) throw new Error('video export: no 2d context')
    ctx.fillStyle = '#000'
    ctx.fillRect(0, 0, this.width, this.height)
    ctx.drawImage(drawn, 0, 0, this.width, this.height)
    return out
  }

  dispose(): void {
    this.root.unmount()
    this.container.remove()
  }
}
