var GasPlugin = require("gas-webpack-plugin");

module.exports = {
    context: __dirname,
    entry: './src/index.js',
    module: {
        rules: [
            {
                test: /\.(js)$/,
                exclude: /node_modules/,
                use: ['babel-loader']
            }
        ]
    },
    resolve: {
        extensions: ['.js']
    },
    output: {
        path: __dirname,
        filename: 'dist/main.gs'
    },
    mode: 'production', // ONLY WORKS IN PRODUCTION ENV
    plugins: [
        new GasPlugin()
    ]
};
