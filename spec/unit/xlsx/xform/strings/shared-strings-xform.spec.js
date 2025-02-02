import fs from 'node:fs';

import testXformHelper from '../test-xform-helper';

import SharedStringsXform from '#lib/xlsx/xform/strings/shared-strings-xform.js';

const expectations = [
  {
    title: 'Shared Strings',
    create() {
      return new SharedStringsXform();
    },
    preparedModel: require('./data/sharedStrings.json'),
    xml: fs.readFileSync(`${__dirname}/data/sharedStrings.xml`).toString(),
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('SharedStringsXform', () => {
  testXformHelper(expectations);
});
