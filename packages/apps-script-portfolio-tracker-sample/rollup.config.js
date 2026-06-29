import { nodeResolve } from "@rollup/plugin-node-resolve";
import typescript from "@rollup/plugin-typescript";

export default {
  input: "src/main.ts",
  output: {
    file: "dist/app.js",
    format: "esm",
  },
  plugins: [
    nodeResolve({
      extensions: [".ts", ".js"],
    }),
    typescript({ tsconfig: "./tsconfig.build.json" }),
  ],
};
