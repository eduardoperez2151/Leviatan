const HtmlWebpackPlugin = require('html-webpack-plugin');

module.exports = {
    devtool: "source-map",
    entry: {
        polyfill: 'babel-polyfill',
        index: './src/js/index.js',
        worksheet:'./src/js/worksheet-helper.js'
    },
    resolve: {
        extensions: [".ts", ".tsx", ".html", ".js"]
    },
    module: {
        rules: [
            {
                test: /\.js$/,
                exclude: /node_modules/,
                use: 'babel-loader'
            },
            {
                test: /\.html$/,
                exclude: /node_modules/,
                use: 'html-loader'
            },
            {
                test: /\.(png|jpg|jpeg|gif|svg)$/,
                use: 'file-loader'
            },
            {
                test: /\.css$/i,
                use: ['style-loader', 'css-loader'],
            }
        ]
    },
    plugins: [
        new HtmlWebpackPlugin({
            filename: "index.html",
            template: './src/pages/index.html',
            chunks: ['polyfill', 'index','worksheet']
        }),
        new HtmlWebpackPlugin({
            filename: "dialog.html",
            template: './src/pages/dialog.html',
            chunks: ['polyfill']
        })

    ]
};