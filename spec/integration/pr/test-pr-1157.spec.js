import ExcelJS from '#lib';
import { getTempFileName } from '../../utils';

describe('github issues', () => {
  it('pull request 1204 - Read and write data validation should be successful', async () => {
    const testFileName = getTempFileName();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('./spec/integration/data/test-pr-1204.xlsx');
    const expected = {
      E1: {
        type: 'textLength',
        formulae: [2],
        showInputMessage: true,
        showErrorMessage: true,
        operator: 'greaterThan',
      },
      E4: {
        type: 'textLength',
        formulae: [2],
        showInputMessage: true,
        showErrorMessage: true,
        operator: 'greaterThan',
      },
    };
    const ws = wb.getWorksheet(1);
    expect(ws.dataValidations.model).to.deep.equal(expected);
    await wb.xlsx.writeFile(testFileName);
  });
});
