const testXformHelper = require('../../test-xform-helper');

const SqrefExtXform = require('#lib/xlsx/xform/sheet/cf-ext/sqref-ext-xform.js');

const expectations = [
  {
    title: 'range',
    create() {
      return new SqrefExtXform();
    },
    preparedModel: 'A1:C3',
    xml: '<xm:sqref>A1:C3</xm:sqref>',
    parsedModel: 'A1:C3',
    tests: ['render', 'parse'],
  },
];

describe('SqrefExtXform', () => {
  testXformHelper(expectations);
});
