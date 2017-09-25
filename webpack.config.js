var path = require("path");
var webpack = require("webpack");
const UglifyJSPlugin = require('uglifyjs-webpack-plugin');
var CopyWebpackPlugin = require('copy-webpack-plugin');

module.exports = {
    target: "web",
    entry: {
        App: "./scripts/App.tsx",
        AllBugBashesView: "./scripts/Components/AllBugBashesView.tsx",
        BugBashEditor: "./scripts/Components/BugBashEditor.tsx",
        BugBashResults: "./scripts/Components/BugBashResults.tsx",
        BugBashCharts: "./scripts/Components/BugBashCharts.tsx",
        SettingsPanel: "./scripts/Components/SettingsPanel.tsx"
    },
    output: {
        filename: "scripts/[name].js",
        libraryTarget: "amd"
    },
    externals: [
        {
            "q": true,
            "react": true,
            "react-dom": true
        },
        /^VSS\/.*/, /^TFS\/.*/, /^q$/
    ],
    resolve: {
        extensions: [".webpack.js", ".web.js", ".ts", ".tsx", ".js"],
        moduleExtensions: ["-loader"],
        alias: { 
            "OfficeFabric": path.resolve(__dirname, "node_modules/office-ui-fabric-react/lib-amd"),
            "MB": path.resolve(__dirname, "node_modules/vsts-extension-react-widgets/lib-amd")
        }        
    },
    module: {
        rules: [
            {
                test: /\.tsx?$/,
                use: "ts-loader"
            },
            {
                test: /\.s?css$/,
                use: [
                    { loader: "style-loader" },
                    { loader: "css-loader" },
                    { loader: "sass-loader" }
                ]
            },
            {
                test: /\.(otf|eot|svg|ttf|woff|woff2)(\?.+)?$/,
                use: "url-loader?limit=4096&name=[name].[ext]"
            }
        ]
    },
    plugins: [
        new webpack.DefinePlugin({
            "process.env.NODE_ENV": JSON.stringify("production")
        }),
        new webpack.optimize.CommonsChunkPlugin({
            name: "common_chunks",
            filename: "./scripts/common_chunks.js",
            minChunks: 3
        }),
        new UglifyJSPlugin({
            minimize: true,
            compress: {
                warnings: false
            },
            output: {
                comments: false
            }
        }),
        new CopyWebpackPlugin([
            { from: "./node_modules/vss-web-extension-sdk/lib/VSS.SDK.min.js", to: "scripts/libs/VSS.SDK.min.js" },
            { from: "./node_modules/es6-promise/dist/es6-promise.min.js", to: "scripts/libs/es6-promise.min.js" },

            { from: "./node_modules/trumbowyg/dist/ui/icons.svg", to: "css/libs/icons.png" },
            
            { from: "./img", to: "img" },
            { from: "./index.html", to: "./" },
            { from: "./README.md", to: "README.md" },
            { from: "./vss-extension.json", to: "vss-extension.json" }
        ])
    ]
}