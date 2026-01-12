import typescript from 'rollup-plugin-typescript2'
import babel from '@rollup/plugin-babel';
import { terser } from 'rollup-plugin-terser'
import { nodeResolve } from '@rollup/plugin-node-resolve';
import commonjs from '@rollup/plugin-commonjs';

export default [
  {
    input: "./src/core/Excel.ts",
    output: [
      {
        file: "./lib/excel.umd.js",
        format: "umd",
        name: 'Excel',
        compact: true,
      },
      {
        file: "./lib/excel.esm.js",
        format: "es",
      },
    ],
    mode: 'production',
    plugins: [
      nodeResolve({
        browser: true, // 针对浏览器环境
        preferBuiltins: false // 不使用Node.js内置模块
      }),
      // 配置commonjs插件来转换CommonJS模块
      commonjs({
        include: 'node_modules/**' // 包含node_modules中的CommonJS模块
      }),
      babel({
        exclude: 'node_modules/**',
        presets: ['@babel/preset-env']
      }),
      typescript({ tsconfig: 'tsconfig.json', declaration: true }),
      // terser(), // 如果需要压缩可以取消注释
    ]
  },
];