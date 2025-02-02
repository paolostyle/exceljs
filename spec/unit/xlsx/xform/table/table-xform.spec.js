import fs from 'node:fs';

import testXformHelper from '../test-xform-helper';

import TableXform from '#lib/xlsx/xform/table/table-xform.js';

const expectations = [
  {
    title: 'showing filter',
    create() {
      return new TableXform();
    },
    initialModel: null,
    preparedModel: require('./data/table.1.1.json'),
    xml: fs.readFileSync(`${__dirname}/data/table.1.2.xml`).toString(),
    parsedModel: require('./data/table.1.3.json'),
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('TableXform', () => {
  testXformHelper(expectations);
});
