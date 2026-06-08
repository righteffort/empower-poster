import { nodeResolve } from "@rollup/plugin-node-resolve";
import typescript from "@rollup/plugin-typescript";

export default {
  input: "src/app.ts",
  output: {
    dir: "dist",
    format: "cjs",
  },
  plugins: [
    nodeResolve({
      extensions: [".ts", ".js"],
    }),
    typescript(),
  ],
};
