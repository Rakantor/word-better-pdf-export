/* eslint-disable no-undef */
const fs = require("fs");
const os = require("os");
const path = require("path");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const urlDev = "https://localhost:3000/";
const urlProd = "https://rakantor.github.io/word-better-pdf-export/"; // GitHub Pages

/**
 * HTTPS certificates for the dev server.
 *  - Windows/macOS: office-addin-dev-certs generates them and trusts the CA in the OS store (what Word needs).
 *  - Linux/WSL: only generate. Word runs on the Windows side, so trust ~/.office-addin-dev-certs/ca.crt there
 *    (see README) instead of prompting for sudo to install it into the Linux store.
 */
async function getHttpsOptions() {
  const dir = path.join(os.homedir(), ".office-addin-dev-certs");
  const files = { ca: path.join(dir, "ca.crt"), cert: path.join(dir, "localhost.crt"), key: path.join(dir, "localhost.key") };
  const exists = Object.values(files).every((f) => fs.existsSync(f));
  if (!exists) {
    if (process.platform === "linux") {
      const { generateCertificates } = require("office-addin-dev-certs/lib/generate");
      await generateCertificates();
      console.log(`Generated dev certificates in ${dir}. Import ${files.ca} into Windows' "Trusted Root Certification Authorities".`);
    } else {
      const devCerts = require("office-addin-dev-certs");
      const opts = await devCerts.getHttpsServerOptions();
      return { ca: opts.ca, key: opts.key, cert: opts.cert };
    }
  }
  return { ca: fs.readFileSync(files.ca), cert: fs.readFileSync(files.cert), key: fs.readFileSync(files.key) };
}

function escapeRegExp(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  return {
    devtool: dev ? "source-map" : false,
    entry: {
      taskpane: "./src/taskpane/taskpane.ts",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].js",
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: { loader: "ts-loader", options: { transpileOnly: true } },
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
        scriptLoading: "blocking",
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "assets/*", to: "assets/[name][ext]" },
          { from: "src/taskpane/taskpane.css", to: "taskpane.css" },
          {
            from: "manifest*.xml",
            to: "[name][ext]",
            // In production, point every localhost URL at the deployed site. <AppDomain> carries the
            // bare origin (no trailing slash), so that is rewritten to the production origin.
            transform(content) {
              if (dev) return content;
              return content
                .toString()
                .replace(new RegExp(escapeRegExp(urlDev), "g"), urlProd)
                .replace(new RegExp(escapeRegExp(urlDev.replace(/\/$/, "")), "g"), new URL(urlProd).origin);
            },
          },
        ],
      }),
    ],
    performance: { hints: false },
    devServer: {
      headers: { "Access-Control-Allow-Origin": "*" },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
      static: false,
    },
  };
};
