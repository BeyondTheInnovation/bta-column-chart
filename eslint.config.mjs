import eslintPluginPowerbiVisuals from "eslint-plugin-powerbi-visuals";

export default [
    eslintPluginPowerbiVisuals.configs.recommended,
    { ignores: ["node_modules/**", "dist/**", ".vscode/**", ".tmp/**"] },
];
