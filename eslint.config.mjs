import officeAddins from "eslint-plugin-office-addins";
import tsParser from "@typescript-eslint/parser";
import globals from "globals";

export default [
  ...officeAddins.configs.recommended,
  {
    plugins: {
      "office-addins": officeAddins,
    },
    languageOptions: {
      parser: tsParser,
    }
  },
  ...[{
    ignores: [
        "src/PowerBI-SPC",
        "src/PowerBI-Funnels",
        "*.js",
        "*.mjs"
    ]
  }],
  ...[{
    languageOptions: {
      globals: {
        ...globals.browser,
        "Office": "readonly",
        "Excel": "readonly"
        }
    }
  }]
];
