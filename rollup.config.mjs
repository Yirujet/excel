import typescript from 'rollup-plugin-typescript2'
import babel from '@rollup/plugin-babel';
import { terser } from 'rollup-plugin-terser'
import postcss from 'rollup-plugin-postcss';
import copy from 'rollup-plugin-copy'

export default [
  {
    external: [/node_modules/],
    input: "./src/core/Excel.ts",
    output: {
      file: "./lib/excel.js",
      format: "umd",
      name: 'Excel',
    },
    plugins: [
      postcss({
        extract: true,
        extract: 'excel.css'
      }),
      babel({
        exclude: 'node_modules/**',
        presets: ['@babel/preset-env']
      }),
      typescript({ tsconfig: 'tsconfig.json', declaration: true }),
      terser(),
      copy({
        targets: [
          { src: ['src/assets/fonts/iconfont.ttf'], dest: 'lib/fonts' },
        ]
      }),
    ]
  },
];
