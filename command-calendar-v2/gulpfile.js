'use strict';

const build = require('@microsoft/sp-build-web');

// Preserve upstream behavior but allow packaging in this environment
if (build.eslintCmd) {
  build.eslintCmd.enabled = false;
}

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

build.initialize(require('gulp'));
