import { useState, type ReactNode } from 'react'
import type { ZoteroCommand } from '../../shared/ipc'
import { useI18n, type StringKey } from '../i18n/locale'
import { useRunCommand } from './ribbon-command'
import {
  IconBook,
  IconCitation,
  IconCitationsPane,
  IconGear,
  IconRefresh,
  IconSearch,
  IconUnlink,
} from './icons'
import { BIG, TabProps } from './ribbon-tabs'

interface ZoteroTabProps extends Pick<TabProps, 'hasDoc' | 'setDropdown'> {
  /** footnotes/endnotes hold Zotero citation fields the bridge cannot see yet */
  zoteroNoteFields?: boolean
  /** the library pane is open */
  zoteroLibraryOpen?: boolean
  /** Refresh on Open is on */
  zoteroAutoRefresh?: boolean
}

// the buttons and their order are the Zotero plugin's own tab in Word
const BUTTONS: Array<{
  command: ZoteroCommand
  label: StringKey
  tip: StringKey
  icon: (size: number) => ReactNode
}> = [
  {
    command: 'addEditCitation',
    label: 'zoteroCitation',
    tip: 'zoteroCitationTip',
    icon: (size) => <IconCitation size={size} />,
  },
  {
    command: 'addEditBibliography',
    label: 'zoteroBibliography',
    tip: 'zoteroBibliographyTip',
    icon: (size) => <IconBook size={size} />,
  },
  {
    command: 'setDocPrefs',
    label: 'zoteroDocumentPreferences',
    tip: 'zoteroDocumentSettingsTip',
    icon: (size) => <IconGear size={size} />,
  },
  {
    command: 'refresh',
    label: 'zoteroRefresh',
    tip: 'zoteroRefreshTip',
    icon: (size) => <IconRefresh size={size} />,
  },
  {
    command: 'removeCodes',
    label: 'zoteroRemoveCodes',
    tip: 'zoteroRemoveCodesTip',
    icon: (size) => <IconUnlink size={size} />,
  },
]

export function ZoteroTab({
  hasDoc,
  setDropdown,
  zoteroNoteFields,
  zoteroLibraryOpen,
  zoteroAutoRefresh,
}: ZoteroTabProps) {
  const { t } = useI18n()
  const runCommand = useRunCommand()
  const [busy, setBusy] = useState<ZoteroCommand | null>(null)

  const runZotero = async (command: ZoteroCommand) => {
    if (busy) return
    setDropdown(() => null)
    // every command lets Zotero rebuild the bibliography from the fields it can see;
    // note citations are not among them yet, so their works would silently drop out
    if (zoteroNoteFields) {
      window.alert(t('zoteroNoteFieldsUnsupported'))
      return
    }
    setBusy(command)
    try {
      const result = await window.desktop.zoteroCommand(command)
      if (!result.ok) {
        console.error('Zotero integration failed', result.error)
        window.alert(
          t(
            result.errorCode === 'connection-refused'
              ? 'zoteroConnectionError'
              : 'zoteroOperationError',
          ),
        )
      }
    } catch (error) {
      console.error('Zotero integration failed', error)
      window.alert(t('zoteroOperationError'))
    } finally {
      setBusy(null)
    }
  }

  return (
    <>
      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className="rb-big"
            data-testid="zotero-quick-cite"
            disabled={!hasDoc}
            data-tip={t('zoteroQuickCiteTip')}
            onClick={() => runCommand('zotero-quick-cite')}
          >
            <span className="rb-big-icon">
              <IconSearch size={BIG} />
            </span>
            <span>{t('zoteroQuickCite')}</span>
          </button>
          <button
            className={`rb-big${zoteroLibraryOpen ? ' active' : ''}`}
            data-testid="zotero-library"
            aria-pressed={!!zoteroLibraryOpen}
            disabled={!hasDoc}
            data-tip={t('zoteroLibraryTip')}
            onClick={() => runCommand('zotero-library')}
          >
            <span className="rb-big-icon">
              <IconCitationsPane size={BIG} />
            </span>
            <span>{t('zoteroLibrary')}</span>
          </button>
        </div>
        <div className="ribbon-group-label">{t('zoteroLibraryGroup')}</div>
      </div>
      <div className="ribbon-group">
        <div className="ribbon-group-items">
          {BUTTONS.map(({ command, label, tip, icon }) => (
            <button
              key={command}
              className="rb-big"
              data-testid={`zotero-${command}`}
              disabled={!hasDoc || busy !== null}
              data-tip={t(tip)}
              // command-exempt: Zotero drives the document over its own integration protocol
              onClick={() => void runZotero(command)}
            >
              <span className="rb-big-icon">{icon(BIG)}</span>
              <span>{t(label)}</span>
            </button>
          ))}
          <div className="rb-col">
            <button
              className={`rb-small${zoteroAutoRefresh ? ' active' : ''}`}
              data-testid="zotero-auto-refresh"
              aria-pressed={!!zoteroAutoRefresh}
              data-tip={t('zoteroAutoRefreshTip')}
              onClick={() => runCommand('zotero-auto-refresh')}
            >
              <IconRefresh size={14} />
              <span className="rb-small-label">{t('zoteroAutoRefresh')}</span>
            </button>
          </div>
        </div>
        <div className="ribbon-group-label">{t('zoteroGroup')}</div>
      </div>
    </>
  )
}
