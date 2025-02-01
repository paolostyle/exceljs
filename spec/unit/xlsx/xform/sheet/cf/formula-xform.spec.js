const testXformHelper = require('../../test-xform-helper');

const FormulaXform = require('#lib/xlsx/xform/sheet/cf/formula-xform.js');

const expectations = [
  {
    title: 'formula',
    create() {
      return new FormulaXform();
    },
    preparedModel: 'ROW()',
    xml: '<formula>ROW()</formula>',
    parsedModel: 'ROW()',
    tests: ['render', 'parse'],
  },
];

describe('FormulaXform', () => {
  testXformHelper(expectations);
});
