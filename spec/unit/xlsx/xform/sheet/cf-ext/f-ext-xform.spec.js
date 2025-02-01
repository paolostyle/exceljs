const testXformHelper = require('../../test-xform-helper');

const FExtXform = require('#lib/xlsx/xform/sheet/cf-ext/f-ext-xform.js');

const expectations = [
  {
    title: 'formula',
    create() {
      return new FExtXform();
    },
    preparedModel: '7',
    xml: '<xm:f>7</xm:f>',
    parsedModel: '7',
    tests: ['render', 'parse'],
  },
];

describe('FExtXform', () => {
  testXformHelper(expectations);
});
