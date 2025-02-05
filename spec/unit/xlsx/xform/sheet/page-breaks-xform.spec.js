import testXformHelper from '../test-xform-helper';

import PageBreaksXform from '#lib/xlsx/xform/sheet/page-breaks-xform.js';

const expectations = [
  {
    title: 'one page break',
    create() {
      return new PageBreaksXform();
    },
    initialModel: { id: 2, max: 3, min: 1, man: 1 },
    preparedModel: { id: 2, max: 3, min: 1, man: 1 },
    xml: '<brk id="2" max="3" min="1" man="1"/>',
    parsedModel: { id: 2, max: 3, min: 1, man: 1 },
    tests: ['prepare', 'render', 'renderIn'],
  },
];

describe('PageBreaksXform', () => {
  testXformHelper(expectations);
});
