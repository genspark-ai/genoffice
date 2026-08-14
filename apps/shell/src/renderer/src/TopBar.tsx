import type { FC } from 'react'
import logoLockup from './assets/genoffice-logo.svg'

export const TopBar: FC = () => {
  const isMac = typeof navigator !== 'undefined' && /mac/i.test(navigator.userAgent || navigator.platform)

  return (
    <div className="tab-bar" style={{ zIndex: 11, borderBottom: '1px solid var(--tab-separator)' }}>
      <div className="tab-bar-drag-spacer" style={{ width: '12px' }} />
      <div className="tab-list" style={{ flexGrow: 1, display: 'flex', alignItems: 'center' }}>
        <img className="logo-lockup" src={logoLockup} alt="GenOffice" style={{ height: '20px', marginRight: '16px', WebkitAppRegion: 'no-drag' } as any} />
        {!isMac && (
          <div style={{ display: 'flex', gap: '4px', WebkitAppRegion: 'no-drag' } as any}>
            <button className="tab-new-btn" style={{ width: 'auto', padding: '0 12px' }} onClick={(e) => { const r = e.currentTarget.getBoundingClientRect(); window.aiOffice?.showAppMenu?.('file', r.left, r.bottom) }}>File</button>
            <button className="tab-new-btn" style={{ width: 'auto', padding: '0 12px' }} onClick={(e) => { const r = e.currentTarget.getBoundingClientRect(); window.aiOffice?.showAppMenu?.('edit', r.left, r.bottom) }}>Edit</button>
            <button className="tab-new-btn" style={{ width: 'auto', padding: '0 12px' }} onClick={(e) => { const r = e.currentTarget.getBoundingClientRect(); window.aiOffice?.showAppMenu?.('window', r.left, r.bottom) }}>Window</button>
            <button className="tab-new-btn" style={{ width: 'auto', padding: '0 12px' }} onClick={(e) => { const r = e.currentTarget.getBoundingClientRect(); window.aiOffice?.showAppMenu?.('help', r.left, r.bottom) }}>Help</button>
          </div>
        )}
      </div>
      {!isMac && (
        <div style={{ display: 'flex', WebkitAppRegion: 'no-drag' } as any}>
          <button className="tab-overflow-btn" title="Minimize" onClick={() => window.aiOffice?.minimizeWindow?.()}>
            <svg width="10" height="10" viewBox="0 0 10 10"><path fill="currentColor" d="M0 4.5v1h10v-1z" /></svg>
          </button>
          <button className="tab-overflow-btn" title="Maximize" onClick={() => window.aiOffice?.maximizeWindow?.()}>
            <svg width="10" height="10" viewBox="0 0 10 10"><path fill="currentColor" d="M1 1v8h8V1H1zm1 1h6v6H2V2z" /></svg>
          </button>
          <button className="tab-overflow-btn" title="Close" onClick={() => window.aiOffice?.closeWindow?.()}>
            <svg width="10" height="10" viewBox="0 0 10 10"><path fill="currentColor" d="M9.4 1.6L8.4.6 5 4 1.6.6.6 1.6 4 5 .6 8.4l1 1L5 6l3.4 3.4 1-1L6 5l3.4-3.4z" /></svg>
          </button>
        </div>
      )}
    </div>
  )
}

export default TopBar
