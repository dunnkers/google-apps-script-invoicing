var GasPlugin = require("gas-webpack-plugin");

module.exports = {
    context: __dirname,
    output: {
        path: __dirname,
        filename: 'dist/main.gs'
    },
    mode: 'production', // ONLY WORKS IN PRODUCTION ENV
    plugins: [
        new GasPlugin()
    ]
};
