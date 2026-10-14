import '@genoffice/ui/tokens.css'
import '../../apps/shell/src/renderer/src/home.css'
import { createRoot } from 'react-dom/client'
import { LocaleProvider } from '../../apps/shell/src/renderer/src/locale'
import { CommandPalette } from '../../apps/shell/src/renderer/src/CommandPalette'

/**
 * The command palette as the shell draws it: the real component, the real
 * stylesheet, and commands shaped exactly as a feature registers them.
 */
const commands = [
  {
    id: 'tidy-headings',
    name: 'Tidy headings',
    description: 'unfold every wrapped heading',
    kind: 'command',
    command: {
      label: 'Tidy headings',
      description: 'unfold every wrapped heading in this report',
      action: { kind: 'ask-assistant', prompt: 'Tidy the headings in this report' },
    },
    createdAt: 'x',
    updatedAt: 'x',
  },
  {
    id: 'open-style-guide',
    name: 'Open style guide',
    description: 'the house style',
    kind: 'command',
    command: {
      label: 'Open the style guide',
      description: 'the folder every report is checked against',
      action: { kind: 'open', path: '/Users/me/Documents/house-style/style.md' },
    },
    createdAt: 'x',
    updatedAt: 'x',
  },
  {
    id: 'file-sidebar',
    name: 'Files',
    description: 'recent files',
    kind: 'command',
    command: {
      label: 'Show recent files',
      action: { kind: 'panel', id: 'file-sidebar' },
    },
    createdAt: 'x',
    updatedAt: 'x',
  },
]

const actions = {
  open: () => {},
  ask: () => {},
  panel: () => {},
}

createRoot(document.getElementById('root')!).render(
  <LocaleProvider initial="en">
    <div className="home">
      <CommandPalette commands={commands} actions={actions} onClose={() => {}} />
    </div>
  </LocaleProvider>,
)
