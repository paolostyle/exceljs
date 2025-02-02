import ExcelJS from '#lib';
import { getTempFileName } from '../../utils/index';

describe('github issues', () => {
  it('issue 234 - Broken XLSX because of "vertical tab" ascii character in a cell', async () => {
    const testFileName = getTempFileName();

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');

    // Start of Heading
    ws.getCell('A1').value = 'Hello, \x01World!';

    // Vertical Tab
    ws.getCell('A2').value = 'Hello, \x0bWorld!';

    await wb.xlsx.writeFile(testFileName);
    const wb2 = new ExcelJS.Workbook();
    const wb2_1 = await wb2.xlsx.readFile(testFileName);
    const ws2 = wb2_1.getWorksheet('Sheet1');

    expect(ws2.getCell('A1').value).to.equal('Hello, World!');
    expect(ws2.getCell('A2').value).to.equal('Hello, World!');
  });
});
