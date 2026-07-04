/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const path = require("path");

const urlDev = "https://localhost:3000/";
const urlProd = "https://www.contoso.com/"; // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const isDev = dev;

  return {
    mode: dev ? "development" : "production",
    devtool: dev ? "eval-source-map" : false,

    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/taskpane/taskpane.ts", "./src/taskpane/taskpane.html"],
      commands: "./src/commands/commands.ts",
    },

    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].js",
      library: { type: "umd" },
      clean: true,
    },

    resolve: { extensions: [".ts", ".html", ".js"] },

    module: {
      rules: [
        { test: /\.ts$/, exclude: /node_modules/, use: [{ loader: "ts-loader", options: { transpileOnly: true } }] },
        { test: /\.html$/, exclude: /node_modules/, use: "html-loader" },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: { filename: "assets/[name][ext][query]" },
        },
        { test: /\.css$/, use: ["style-loader", "css-loader"] },
      ],
    },

    plugins: [
      new HtmlWebpackPlugin({
        template: "src/taskpane/taskpane.html",
        filename: "taskpane.html",
        chunks: ["polyfill", "taskpane"],
        cache: false,
      }),
      new HtmlWebpackPlugin({
        template: "src/commands/commands.html",
        filename: "commands.html",
        chunks: ["polyfill", "commands"],
        cache: false,
        inject: "body",
        hash: true,
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets",
            to: "assets",
            globOptions: { ignore: ["**/generate-icons.js", "**/generate-icons.js.map"] },
          },
          {
            from: "manifest*.xml",
            to: "[name][ext]",
            transform(content) {
              if (dev) return content;
              return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
            },
          },
        ],
      }),
    ],

    devServer: {
      headers: { "Access-Control-Allow-Origin": "*" },
      server: {
        type: "https",
        options:
          env.WEBPACK_BUILD || options.https !== undefined
            ? options.https
            : await getHttpsOptions(),
      },
      port: 3000,
      hot: true,
    },
  };
};
