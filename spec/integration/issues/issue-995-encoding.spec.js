import ExcelJS from '#lib';

const TEST_CSV_FILE_NAME = './spec/out/issue-995-encoding.test.csv';
const HEBREW_TEST_STRING = 'משהו שכתוב בעברית';

describe('github issues', () => {
  it('issue 995 - encoding option works fine', { timeout: 6000 }, async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('wheee');
    ws.getCell('A1').value = HEBREW_TEST_STRING;

    const options = {
      encoding: 'UTF-8',
    };
    await wb.csv.writeFile(TEST_CSV_FILE_NAME, options);
    const ws2 = new ExcelJS.Workbook();
    const ws2_1 = await ws2.csv.readFile(TEST_CSV_FILE_NAME, options);
    expect(ws2_1.getCell('A1').value).to.equal(HEBREW_TEST_STRING);
  });
});
