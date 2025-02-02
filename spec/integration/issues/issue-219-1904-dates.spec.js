import ExcelJS from '#lib';
import { getTempFileName } from '../../utils/index';

describe('github issues', () => {
  describe('issue 219 - 1904 dates not supported', () => {
    it('Reading 1904.xlsx', async () => {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.readFile('./spec/integration/data/1904.xlsx');

      expect(wb.properties.date1904).to.equal(true);

      const ws = wb.getWorksheet('Sheet1');

      expect(ws.getCell('B4').value.toISOString()).to.equal(
        '1904-01-01T00:00:00.000Z',
      );
    });

    it('Writing and Reading', async () => {
      const testFileName = getTempFileName();

      const wb = new ExcelJS.Workbook();
      wb.properties.date1904 = true;

      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('B4').value = new Date('1904-01-01T00:00:00.000Z');

      await wb.xlsx.writeFile(testFileName);

      const wb2 = new ExcelJS.Workbook();
      const wb2_1 = await wb2.xlsx.readFile(testFileName);

      expect(wb2_1.properties.date1904).to.equal(true);

      const ws2 = wb2_1.getWorksheet('Sheet1');

      expect(ws2.getCell('B4').value.toISOString()).to.equal(
        '1904-01-01T00:00:00.000Z',
      );
    });
  });
});
