export default {
  external: [/node_modules/],
  input: "./src/main.ts",
  output: {
    file: "./dist/bundle.js",
    format: "umd",
    name: 'Excel'
  },
};
