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
    taskpane: "./src/taskpane.ts" // تأكد أنه يستخدم TypeScript
  },
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "[name].js",
    clean: true
  },
  resolve: {
    extensions: [".ts", ".js"]
  },
  module: {
    rules: [
      {
        test: /\.ts$/,
        include: [path.resolve(__dirname, "src")],
        use: [
          {
            loader: "ts-loader",
            options: { transpileOnly: true } // فقط الترجمة بدون التحقق الكامل
          }
        ]
      },
      { test: /\.css$/, use: ["style-loader", "css-loader"] },
      { test: /\.(png|jpg|gif|svg)$/, type: "asset/resource" },
      { test: /\.html$/, use: "html-loader" }
    ]
  },
  devServer: {
    static: [
      { directory: path.join(__dirname, "dist"), serveIndex: false },
      { directory: path.join(__dirname, "assets"), serveIndex: false },
      { directory: path.join(__dirname, "examples"), serveIndex: false }
    ],
    server: { type: "https", options: httpsConfig },
    host: "localhost",
    port: 3000,
    hot: true,
    allowedHosts: "all",
    headers: { "Access-Control-Allow-Origin": "*" },
    historyApiFallback: {
      index: "/taskpane.html",
      rewrites: [{ from: /^\/$/, to: "/taskpane.html" }]
    }
  }
};
