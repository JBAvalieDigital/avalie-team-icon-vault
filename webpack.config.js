const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

module.exports = {
  entry: "./src/taskpane/taskpane.ts",
  output: {
    path: path.resolve(__dirname, "docs"),
    filename: "taskpane.js",
    clean: true,
  },
  resolve: {
    extensions: [".ts", ".js"],
  },
  module: {
    rules: [
      {
        test: /\.ts$/,
        use: "ts-loader",
        exclude: /node_modules/,
      },
      {
        test: /\.css$/,
        use: ["style-loader", "css-loader"],
      },
    ],
  },
  plugins: [
    new HtmlWebpackPlugin({
      template: "./src/taskpane/taskpane.html",
      filename: "taskpane.html",
    }),
    new CopyWebpackPlugin({
      patterns: [
        // Icon JSON data files — fetched lazily at runtime
        { from: "node_modules/@iconify-json/fluent/icons.json", to: "icons/fluent.json" },
        { from: "node_modules/@iconify-json/lucide/icons.json", to: "icons/lucide.json" },
        { from: "node_modules/@iconify-json/tabler/icons.json", to: "icons/tabler.json" },
        { from: "node_modules/@iconify-json/heroicons/icons.json", to: "icons/heroicons.json" },
        { from: "node_modules/@iconify-json/ri/icons.json", to: "icons/ri.json" },
        { from: "node_modules/@iconify-json/fa6-regular/icons.json", to: "icons/fa6-regular.json" },
        { from: "node_modules/@iconify-json/fa6-solid/icons.json", to: "icons/fa6-solid.json" },
        // Ribbon icon assets
        { from: "assets", to: "assets" },
      ],
    }),
  ],
  devServer: {
    static: path.resolve(__dirname, "docs"),
    port: 3000,
    server: {
      type: "https",
      options: {
        key: require("os").homedir() + "/.office-addin-dev-certs/localhost.key",
        cert: require("os").homedir() + "/.office-addin-dev-certs/localhost.crt",
      },
    },
    headers: {
      "Access-Control-Allow-Origin": "*",
    },
  },
};
