module.exports = [
  {
    files: ["src/**/*.js", "tests/**/*.js", "scripts/**/*.js", "*.js"],
    languageOptions: {
      ecmaVersion: "latest",
      sourceType: "commonjs",
      globals: { CustomFunctions: "readonly", module: "readonly", require: "readonly", process: "readonly", console: "readonly", __dirname: "readonly", describe: "readonly", test: "readonly", expect: "readonly" }
    },
    rules: { "no-undef": "error", "no-var": "error", "prefer-const": "error" }
  }
];
