const testXformHelper = require('../test-xform-helper');

const BlipXform = require('#lib/xlsx/xform/drawing/blip-xform.js');

const expectations = [
  {
    title: 'full',
    create() {
      return new BlipXform();
    },
    preparedModel: { rId: 'rId1' },
    xml: '<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1" cstate="print" />',
    parsedModel: { rId: 'rId1' },
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('BlipXform', () => {
  testXformHelper(expectations);
});
