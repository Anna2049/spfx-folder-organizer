'use strict';

require('dotenv').config();
const fs = require('fs');
const path = require('path');
const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be made available via the local class names object passed to the 'require' call of the battery. In order to refer to this class in your module, you must use the global CSS class name directly.`);

/* ---------- Inject SHAREPOINT_SITE_URL from .env into serve.json ---------- */
const siteUrl = process.env.SHAREPOINT_SITE_URL;
if (siteUrl) {
  const serveConfigPath = path.join(__dirname, 'config', 'serve.json');
  const serveConfig = JSON.parse(fs.readFileSync(serveConfigPath, 'utf8'));
  serveConfig.initialPage = siteUrl.replace(/\/+$/, '') + '/_layouts/workbench.aspx';
  fs.writeFileSync(serveConfigPath, JSON.stringify(serveConfig, null, 2) + '\n', 'utf8');
}

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);
  result.set('serve', result.get('serve-deprecated'));
  return result;
};

build.initialize(require('gulp'));
