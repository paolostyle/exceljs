import fs from 'node:fs';

import testXformHelper from '../test-xform-helper';

import XmlStream from '#lib/utils/xml-stream.js';
import StylesXform from '#lib/xlsx/xform/style/styles-xform.js';

const expectations = [
  {
    title: 'Styles with fonts',
    create() {
      return new StylesXform();
    },
    preparedModel: require('./data/styles.1.1.json'),
    xml: fs.readFileSync(`${__dirname}/data/styles.1.2.xml`).toString(),
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('StylesXform', () => {
  testXformHelper(expectations);

  describe('As StyleManager', () => {
    it('Renders empty model', () => {
      const stylesXform = new StylesXform(true);
      const expectedXml = fs
        .readFileSync(`${__dirname}/data/styles.2.2.xml`)
        .toString();

      const xmlStream = new XmlStream();
      stylesXform.render(xmlStream);

      expect(xmlStream.xml).xml.to.equal(expectedXml);
    });
  });
});
