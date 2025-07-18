import typescript from 'rollup-plugin-typescript2'
import babel from '@rollup/plugin-babel';
import { terser } from 'rollup-plugin-terser'
import postcss from 'rollup-plugin-postcss';
import copy from 'rollup-plugin-copy'
import autoprefixer from "autoprefixer";
import cssnano from "cssnano";

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
      copy({
        targets: [
          { src: ['src/assets/fonts/iconfont.ttf'], dest: 'lib/fonts' },
          { src: ['src/assets/css/lu2.css'], dest: 'lib/css' },
          { src: ['src/assets/js/lu2.js'], dest: 'lib/js' },
        ]
      }),
    ]
  },
];
