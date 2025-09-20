import { fileURLToPath } from "node:url";
import powerbiVisualsConfigs from "eslint-plugin-powerbi-visuals";

const tsconfigRootDir = fileURLToPath(new URL(".", import.meta.url));
const recommended = powerbiVisualsConfigs.configs.recommended;
const recommendedWithRoot = {
    ...recommended,
    languageOptions: {
        ...(recommended.languageOptions ?? {}),
        parserOptions: {
            ...(recommended.languageOptions?.parserOptions ?? {}),
            tsconfigRootDir,
        },
    },
};

export default [
    {
        ignores: ["node_modules/**", "dist/**", ".vscode/**", ".tmp/**"],
    },
    recommendedWithRoot,
    {
        languageOptions: {
            parserOptions: {
                tsconfigRootDir,
            },
        },
    },
];
