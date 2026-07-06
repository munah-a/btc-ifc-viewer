import eslint from '@eslint/js';
import eslintConfigPrettier from 'eslint-config-prettier';
import tseslint from 'typescript-eslint';

export default tseslint.config(
  {
    ignores: [
      'dist/**',
      'node_modules/**',
      'BIM Viewer UIUX Branding-handoff/**',
      'test-results/**',
      'playwright-report/**',
      'public/**',
      '.claude/**',
    ],
  },
  eslint.configs.recommended,
  tseslint.configs.recommendedTypeChecked,
  {
    languageOptions: {
      parserOptions: {
        project: ['./tsconfig.json', './tsconfig.node.json'],
        tsconfigRootDir: import.meta.dirname,
      },
    },
    rules: {
      // Mirror tsc's noUnusedLocals/noUnusedParameters convention: a leading
      // underscore marks intentionally-unused identifiers.
      '@typescript-eslint/no-unused-vars': [
        'error',
        {
          args: 'all',
          argsIgnorePattern: '^_',
          caughtErrorsIgnorePattern: '^_',
          varsIgnorePattern: '^_',
        },
      ],
    },
  },
  {
    // Plain JS/MJS files (configs, build scripts) — no type information
    // available; they run under Node.
    files: ['**/*.js', '**/*.mjs'],
    extends: [tseslint.configs.disableTypeChecked],
    languageOptions: {
      globals: {
        console: 'readonly',
        process: 'readonly',
        fetch: 'readonly',
        Buffer: 'readonly',
      },
    },
  },
  {
    // src/viewer.ts: the ThatOpen boundary is untyped by design until W2.2
    // extracts core/fragments-model.ts (AUDIT A8). e2e/: specs reach into raw
    // window.__viewer internals until W2.5 ships the frozen, typed
    // window.__viewerTestApi (AUDIT T6). Unsafe-any rules would demand that
    // typing work now, out of W0 scope — re-enable per file as A8/T6 land.
    files: ['src/viewer.ts', 'e2e/**/*.ts'],
    rules: {
      '@typescript-eslint/no-explicit-any': 'off',
      '@typescript-eslint/no-unsafe-argument': 'off',
      '@typescript-eslint/no-unsafe-assignment': 'off',
      '@typescript-eslint/no-unsafe-call': 'off',
      '@typescript-eslint/no-unsafe-member-access': 'off',
      '@typescript-eslint/no-unsafe-return': 'off',
    },
  },
  eslintConfigPrettier,
);
