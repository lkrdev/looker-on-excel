const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const devCerts = require("office-addin-dev-certs");

module.exports = async (env, options) => {
  const isDev = options.mode !== "production";
  let httpsOptions = true;

  if (isDev) {
    try {
      httpsOptions = await devCerts.getHttpsServerOptions();
    } catch (e) {
      console.warn("Using default self-signed HTTPS options", e);
      httpsOptions = true;
    }
  }

  return {
    devtool: isDev ? "source-map" : false,
    entry: {
      taskpane: "./src/taskpane/index.tsx",
      "dialog-auth": "./src/auth/dialog-auth.ts",
      "dialog-callback": "./src/auth/dialog-callback.ts",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].bundle.js",
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          use: "ts-loader",
          exclude: /node_modules/,
        },
        {
          test: /\.css$/,
          use: ["style-loader", "css-loader"],
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico|svg)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext]",
          },
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/index.html",
        chunks: ["taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "dialog-auth.html",
        template: "./src/auth/dialog-auth.html",
        chunks: ["dialog-auth"],
      }),
      new HtmlWebpackPlugin({
        filename: "dialog-callback.html",
        template: "./src/auth/dialog-callback.html",
        chunks: ["dialog-callback"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext]",
            noErrorOnMissing: true,
          },
          {
            from: "manifest.xml",
            to: "manifest.xml",
          },
        ],
      }),
    ],
    devServer: {
      static: {
        directory: path.join(__dirname, "dist"),
      },
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: httpsOptions,
      },
      port: 3000,
      hot: true,
    },
  };
};
