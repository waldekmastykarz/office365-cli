const path = require('path');
const nodeExternals = require('webpack-node-externals');
const UglifyJsPlugin = require('uglifyjs-webpack-plugin');

module.exports = {
  entry: './src/index.ts',
  module: {
    rules: [
      {
        test: /\.tsx?$/,
        use: 'ts-loader',
        exclude: /node_modules/
      }
    ]
  },
  target: 'node',
  externals: [nodeExternals()],
  resolve: {
    extensions: [ '.tsx', '.ts', '.js' ]
  },
  output: {
    filename: 'index.js',
    path: path.resolve(__dirname, 'dist')
  },
  optimization: {
    minimizer: [new UglifyJsPlugin()]
  },
  mode: 'production'
};