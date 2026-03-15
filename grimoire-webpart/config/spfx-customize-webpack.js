/**
 * SPFx 1.22.1 Webpack Customization
 *
 * Adds watchOptions to reduce watcher noise during build-watch.
 */
module.exports = function customize(webpackConfig) {
  // ─── Watch options ────────────────────────────────────────
  if (!webpackConfig.watchOptions) {
    webpackConfig.watchOptions = {};
  }
  webpackConfig.watchOptions.ignored = [
    '**/node_modules/**',
    '**/lib/**',
    '**/lib-commonjs/**',
    '**/dist/**',
    '**/temp/**',
    '**/release/**'
  ];
  webpackConfig.watchOptions.aggregateTimeout = 200;

  return webpackConfig;
};
