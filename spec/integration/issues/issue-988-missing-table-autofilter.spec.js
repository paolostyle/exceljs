import ExcelJS from '#lib';

describe('github issues', () => {
  it('issue 988 - table without autofilter model', () => {
    const wb = new ExcelJS.Workbook();
    wb.xlsx.readFile('./spec/integration/data/test-issue-988.xlsx');
  });
});
