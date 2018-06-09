module.export = {
    module: {
        rules: [{
            exclude: /node_modules/,
            test: /\.js$/,
            use: { loader: 'babel-loader' },
        }],
    },
    output: {
        libraryTarget: 'this'
    }
};
