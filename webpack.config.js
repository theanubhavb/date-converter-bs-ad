const CopyPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CustomFunctionsMetadataPlugin = require("custom-functions-metadata-plugin");
const { getHttpsServerOptions } = require("office-addin-dev-certs");

module.exports = async () => ({
  entry: { functions: "./src/functions/functions.js" },
  devtool: "source-map",
  output: { clean: true, filename: "[name].js" },
  plugins: [
    new HtmlWebpackPlugin({ template: "./src/functions/functions.html", filename: "functions.html", chunks: ["functions"] }),
    new CustomFunctionsMetadataPlugin({ output: "functions.json", input: "./src/functions/functions.js" }),
    new CopyPlugin({ patterns: [{ from: "./src/taskpane/taskpane.html", to: "taskpane.html" }, { from: "./assets", to: "assets" }] })
  ],
  devServer: { server: { type: "https", options: await getHttpsServerOptions() }, port: 3000, static: "./dist", headers: { "Access-Control-Allow-Origin": "*" } }
});
