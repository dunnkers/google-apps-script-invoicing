import GasPlugin from 'gas-webpack-plugin';

export const context = __dirname;
export const entry = './src/index.js';
export const module = {
    rules: [
        {
            test: /\.(js)$/,
            exclude: /node_modules/,
            use: ['babel-loader']
        }
    ]
};
export const resolve = {
    extensions: ['.js']
};
export const output = {
    path: __dirname,
    filename: 'dist/main.gs'
};
export const mode = 'production';
export const plugins = [
    new GasPlugin()
];
