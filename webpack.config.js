const path = require("path");
const CopyWebpackPlugin = require("copy-webpack-plugin");

module.exports = {
  entry: { hub: "./src/hub.js" },
  output: { path: path.resolve(__dirname, "dist"), filename: "[name].js", clean: true },
  resolve: { extensions: [".js"] },
  performance: { hints: false },
  plugins: [new CopyWebpackPlugin({ patterns: [{ from: "src/hub.html", to: "hub.html" }] })]
};
