// Repo-wide lint config (lint only — no formatter, so no whole-repo format diff).
// Rules are tuned so the current codebase passes cleanly; tighten incrementally.
import js from '@eslint/js'
import globals from 'globals'
import tseslint from 'typescript-eslint'
import reactHooks from 'eslint-plugin-react-hooks'

export default tseslint.config(
  {
    ignores: [
      '**/node_modules/**',
      '**/dist/**',
      '**/out/**',
      '**/.build/**',
      '**/release/**',
      '**/target/**',
      '**/coverage/**',
      'scripts/drivers/**',
      'apps/*/build/**',
      'packages/*/src/vendor/**',
    ],
  },
  js.configs.recommended,
  ...tseslint.configs.recommended,
  {
    rules: {
      // The codebase interoperates with untyped vendor APIs (Univer, pptxgenjs,
      // Electron IPC payloads); `any` at those boundaries is deliberate.
      '@typescript-eslint/no-explicit-any': 'off',
      '@typescript-eslint/no-unused-vars': [
        'error',
        {
          argsIgnorePattern: '^_',
          varsIgnorePattern: '^_',
          caughtErrors: 'none',
        },
      ],
      // Empty catch = deliberate fail-open; other empty blocks still flagged.
      'no-empty': ['error', { allowEmptyCatch: true }],
    },
  },
  {
    // tsc already checks undefined identifiers with full type info; eslint's
    // no-undef false-positives on TS-only constructs.
    files: ['**/*.{ts,tsx}'],
    rules: {
      'no-undef': 'off',
    },
  },
  {
    // Classic hooks rules only; the React-Compiler rule set (refs, purity,
    // immutability, …) is not adopted yet.
    files: ['**/*.tsx'],
    plugins: { 'react-hooks': reactHooks },
    rules: {
      'react-hooks/rules-of-hooks': 'error',
      'react-hooks/exhaustive-deps': 'warn',
    },
  },
  {
    // Plain JS build/tool scripts run in Node without type info.
    files: ['**/*.{js,mjs,cjs}'],
    languageOptions: {
      globals: { ...globals.node },
    },
    rules: {
      '@typescript-eslint/no-require-imports': 'off',
    },
  },
)
