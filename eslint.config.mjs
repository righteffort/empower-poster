import eslint from '@eslint/js';
import { defineConfig } from 'eslint/config';
import tseslint from 'typescript-eslint';

export default defineConfig(
  { ignores: ["dist/"] },
  eslint.configs.recommended,
  ...tseslint.configs.strict,
  ...tseslint.configs.stylistic,
    {
        files: ["**/*.ts", "**/*.tsx"],
        rules: {
            "no-unused-vars": "off",
            "@typescript-eslint/no-unused-vars": "off",
        }
    }
);
