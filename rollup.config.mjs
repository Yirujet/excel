import typescript from 'rollup-plugin-typescript2'
import babel from '@rollup/plugin-babel';
import { terser } from 'rollup-plugin-terser'
import postcss from 'rollup-plugin-postcss';
import autoprefixer from "autoprefixer";
import cssnano from "cssnano";
import { nodeResolve } from '@rollup/plugin-node-resolve';
import commonjs from '@rollup/plugin-commonjs';

export default [
  {
    external: [/node_modules/],
    input: "./src/core/Excel.ts",
    output: {
      file: "./lib/excel.js",
      format: "umd",
      // format: 'es',
      name: 'Excel',
      compact: true,
    },
    // mode: 'production',
    plugins: [
      // commonjs(
      //   {
      //     include: 'node_modules/**',
      //   }
      // ),
      postcss({
        extract: true,
        extract: 'excel.css',
        minimize: true,
        plugins: [
          autoprefixer(),
          cssnano()
        ]
      }),
      babel({
        exclude: 'node_modules/**',
        presets: ['@babel/preset-env']
      }),
      typescript({ tsconfig: 'tsconfig.json', declaration: true }),
      // terser(),
      // nodeResolve(),
    ]
  },
];
