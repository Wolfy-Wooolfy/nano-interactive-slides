const path = require("path");
const fs = require("fs");

const userHome = process.env.USERPROFILE || process.env.HOME || "";
const certBase = path.join(userHome, ".office-addin-dev-certs");
const httpsConfig = {
  key: fs.readFileSync(path.join(certBase, "localhost.key")),
  cert: fs.readFileSync(path.join(certBase, "localhost.crt"))
};

module.exports = {
  mode: "development",
  entry: {
    taskpane: "./src/taskpane.tsx"
  },
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "[name].js",
    clean: true
  },
  resolve: {
    extensions: [".ts", ".tsx", ".js"]
  },
  module: {
    rules: [
      { test: /\.tsx?$/, use: "ts-loader", exclude: /node_modules/ },
      { test: /\.css$/, use: ["style-loader", "css-loader"] },
      { test: /\.(png|jpg|gif|svg)$/, type: "asset/resource" },
      { test: /\.html$/, use: "html-loader" }
    ]
  },
  devServer: {
    static: [
      { directory: path.join(__dirname, "dist") },
      { directory: path.join(__dirname, "assets") },
      { directory: path.join(__dirname, "examples") }
    ],
    server: {
      type: "https",
      options: httpsConfig
    },
    host: "localhost",
    port: 3000,
    hot: true,
    allowedHosts: "all",
    headers: {
      "Access-Control-Allow-Origin": "*"
    },
    historyApiFallback: {
      rewrites: [{ from: /^\/taskpane.html$/, to: "/taskpane.html" }]
    }
  }
};
