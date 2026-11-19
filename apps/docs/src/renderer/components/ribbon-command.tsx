import { createContext, useContext, type ReactNode } from 'react'
import {
  COMMANDS,
  type ActionCommandId,
  type CommandDef,
  type CommandId,
} from '../../shared/commands'
import { offByDocument, type DocumentStateInput } from '../commands/menu-state'
import {
  ribbonCommandPressed,
  ribbonControl,
  type RibbonCommandId,
  type RibbonViewState,
} from '../commands/ribbon-controls'
import { useI18n } from '../i18n/locale'
import type { RibbonFormatState } from './ribbon-format-state'

/** what a ribbon control reads to stand for a command */
export interface RibbonCommandSource {
  hasDoc: boolean
  format: RibbonFormatState
  view: RibbonViewState
  run(id: ActionCommandId, payload?: string): void
}

export const RibbonCommandContext = createContext<RibbonCommandSource | null>(null)

/** whether the document and the caret rule a command out; the same rules grey out its menu item */
function useCommandOff(
  id: CommandId,
  over?: Partial<DocumentStateInput>,
): { source: RibbonCommandSource; disabled: boolean } {
  const source = useContext(RibbonCommandContext)
  if (!source) throw new Error('a ribbon command control outside the ribbon')
  const { format, view } = source
  return {
    source,
    disabled: offByDocument(id, {
      hasDoc: source.hasDoc,
      editable: format.editable,
      inTable: format.inTable,
      inTextBox: !!format.sub,
      cellsSelected: format.canMergeCells,
      labelSheet: format.labelSheet,
      hasSelection: format.hasSelection,
      pictureSelected: format.imageSelected,
      hasNotes: format.hasNotes,
      hasChart: format.hasChart,
      canSplitTable: format.canSplitTable,
      canComment: format.canComment,
      inEmptyParagraph: format.inEmptyParagraph,
      hasComments: view.hasComments,
      hasTableOfFigures: format.hasTableOfFigures,
      hasIndex: format.hasIndex,
      hasTableOfAuthorities: format.hasTableOfAuthorities,
      hasMergeSource: view.hasMergeSource,
      windowsDetached: view.windowsDetached,
      ...over,
    }),
  }
}

/**
 * Whether a command's rule greys a control that is not an RbCommand. `over`
 * sets aside an input the control answers for itself: New Comment stays live
 * in a document protected for comments only.
 */
export const useCommandGreyed = (id: CommandId, over?: Partial<DocumentStateInput>): boolean =>
  useCommandOff(id, over).disabled

/** a command as one control shows it */
export function useRibbonCommand(id: RibbonCommandId) {
  const { source, disabled } = useCommandOff(id)
  return {
    disabled,
    pressed: ribbonCommandPressed(id, source.format, source.view),
    control: ribbonControl(id),
    // a placeholder is greyed out and has nothing to run
    run: () => {
      if (!(COMMANDS[id] as CommandDef).planned) source.run(id as ActionCommandId)
    },
  }
}

/** runs any command from inside the ribbon, for the items of a dropdown or a gallery */
export function useRunCommand(): RibbonCommandSource['run'] {
  const source = useContext(RibbonCommandContext)
  if (!source) throw new Error('a ribbon command control outside the ribbon')
  return source.run
}

/**
 * A button that stands for a command. `icon` shows the icon alone; `small` and
 * `big` add the command's caption beside or under it.
 */
export function RbCommand({
  cmd,
  look = 'icon',
  children,
}: {
  cmd: RibbonCommandId
  look?: 'icon' | 'small' | 'big'
  children: ReactNode
}) {
  const { t } = useI18n()
  const command = useRibbonCommand(cmd)
  const tip = t(command.control.tip)
  const caption = command.control.caption ? t(command.control.caption) : ''
  if (look === 'big') {
    return (
      <button
        className={`rb-big${command.pressed ? ' active' : ''}`}
        disabled={command.disabled}
        data-tip={tip}
        onClick={command.run}
      >
        <span className="rb-big-icon">{children}</span>
        <span>{caption}</span>
      </button>
    )
  }
  if (look === 'small') {
    return (
      <button
        className={`rb-small${command.pressed ? ' active' : ''}`}
        disabled={command.disabled}
        data-tip={tip}
        onClick={command.run}
      >
        {children}
        <span className="rb-small-label">{caption}</span>
      </button>
    )
  }
  return (
    <button
      className={`rb-icon ${command.pressed ? 'active' : ''}`}
      disabled={command.disabled}
      data-tip={tip}
      aria-label={tip}
      onClick={command.run}
    >
      {children}
    </button>
  )
}

/** a dropdown item that stands for a command: greyed by the command's own rule */
export function RbMenuCommand({
  cmd,
  onRun,
  children,
}: {
  cmd: ActionCommandId
  onRun: () => void
  children: ReactNode
}) {
  const { source, disabled } = useCommandOff(cmd)
  const command = {
    run: () => {
      source.run(cmd)
      onRun()
    },
  }
  return (
    <button disabled={disabled} onClick={command.run}>
      {children}
    </button>
  )
}
