const path = require('path');

module.exports = {
  mode: 'production',
  entry: './src/main.js',
  output: {
    path: path.resolve(__dirname, 'dist/assets'),
    filename: 'app.js',
    clean: true,
  },
  devtool: false,
};
