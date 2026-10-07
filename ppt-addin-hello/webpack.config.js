/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

// One port, set in package.json (config.dev_server_port); `npm run set-port`
// rewrites it there and in both manifests.
const port = process.env.npm_package_config_dev_server_port || 3000;
const urlDev = "https://localhost:" + port + "/";
// Dev only for this spike: production hosting is decided later.
const urlProd = urlDev;

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./taskpane.html"],
    },
    output: {
      clean: true,
    },
    resolve: {
      extensions: [".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader"
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.json",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
          },
        ],
      }),
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
      // The pane is served over https, so it cannot call the http backend
      // directly (mixed content). Proxying keeps every call same-origin.
      // Backend: ppt_gen_v3p1, run with uvicorn on 127.0.0.1:8000.
      proxy: [
        {
          context: ["/api", "/download"],
          target: process.env.PB_BACKEND || "http://127.0.0.1:8000",
        },
      ],
      // office.js runs cross-origin, so its errors reach the page as a bare
      // "Script error." and the overlay covers the pane. Errors go to the log.
      client: {
        overlay: { runtimeErrors: false },
      },
    },
  };

  return config;
};
