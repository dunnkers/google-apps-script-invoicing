var webpack = require('webpack');

module.exports = {
    output: {
        libraryTarget: 'this',
        library: 'henky'
    },
    mode: 'development',
    node: false,
    plugins: [
        new webpack.ProvidePlugin({
            dollo: 'util'
        })
    ]
};
