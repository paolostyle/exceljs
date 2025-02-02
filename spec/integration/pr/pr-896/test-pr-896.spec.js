import Excel from '#lib';
import { getTempFileName } from '../../../utils/index';
import tools from '../../../utils/tools';

const RT_ARR = [
  { text: 'First Line:\n', font: { bold: true } },
  { text: 'Second Line\n' },
  { text: 'Third Line\n' },
  { text: 'Last Line' },
];
const TEST_VALUE = {
  richText: RT_ARR,
};
const TEST_NOTE = {
  texts: RT_ARR,
};

describe('pr related issues', () => {
  describe('pr 896 add xml:space="preserve" for all whitespaces', () => {
    it('should store cell text and comment with leading new line', () => {
      const testFileName = getTempFileName();

      const properties = tools.fix(
        require('../../../utils/data/sheet-properties.json'),
      );
      const pageSetup = tools.fix(
        require('../../../utils/data/page-setup.json'),
      );

      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('sheet1', {
        properties,
        pageSetup,
      });

      ws.getColumn(1).width = 20;
      ws.getCell('A1').value = TEST_VALUE;
      ws.getCell('A1').note = TEST_NOTE;
      ws.getCell('A1').alignment = { wrapText: true };

      return wb.xlsx
        .writeFile(testFileName)
        .then(() => {
          const wb2 = new Excel.Workbook();
          return wb2.xlsx.readFile(testFileName);
        })
        .then((wb2) => {
          const ws2 = wb2.getWorksheet('sheet1');
          expect(ws2).not.toBeUndefined();
          expect(ws2.getCell('A1').value).to.deep.equal(TEST_VALUE);
        });
    });
  });
});
