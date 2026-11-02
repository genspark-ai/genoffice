import { Component, type ReactNode } from 'react'

/**
 * One broken overlay layer must not blank the whole viewer: a render-phase
 * exception inside whatever this wraps degrades to the fallback (nothing, by
 * default) while the rest of the app keeps working. genoffice#2424 — a page
 * overlay drawing bad geometry took the entire window down with it.
 */
export class LayerBoundary extends Component<
  { children: ReactNode; fallback?: ReactNode },
  { failed: boolean }
> {
  state = { failed: false }

  static getDerivedStateFromError(): { failed: boolean } {
    return { failed: true }
  }

  componentDidCatch(error: unknown): void {
    // the layer is degraded, not the app: keep the reason on the console so a
    // report can name the layer instead of a blank window
    console.error('[pdf] an overlay layer failed and was skipped:', error)
  }

  render(): ReactNode {
    return this.state.failed ? (this.props.fallback ?? null) : this.props.children
  }
}
