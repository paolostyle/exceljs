import fs from 'node:fs';

import testXformHelper from '../test-xform-helper';

import WorkbookXform from '#lib/xlsx/xform/book/workbook-xform.js';

const expectations = [
  {
    title: 'book.1',
    create() {
      return new WorkbookXform();
    },
    preparedModel: require('./data/book.1.1.json'),
    xml: fs
      .readFileSync(`${__dirname}/data/book.1.2.xml`)
      .toString()
      .replace(/\r\n/g, '\n'),
    parsedModel: require('./data/book.1.3.json'),
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'book.2 - no properties',
    create() {
      return new WorkbookXform();
    },
    xml: fs
      .readFileSync(`${__dirname}/data/book.2.2.xml`)
      .toString()
      .replace(/\r\n/g, '\n'),
    parsedModel: require('./data/book.2.3.json'),
    tests: ['parse'],
  },
];

describe('WorkbookXform', () => {
  testXformHelper(expectations);
});
