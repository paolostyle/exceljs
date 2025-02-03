import ExcelJS from '#lib';

describe('ExcelJS', () => {
  it('should read and write xlsx via binary buffer', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('blort');

    ws.getCell('A1').value = 'Hello, World!';
    ws.getCell('A2').value = 7;

    const buffer = await wb.xlsx.writeBuffer();
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(buffer);
    const ws2 = wb2.getWorksheet('blort');

    expect(ws2.getCell('A1').value).to.equal('Hello, World!');
    expect(ws2.getCell('A2').value).to.equal(7);
  });
  it('should read and write xlsx via base64 buffer', async () => {
    const options = {
      base64: true,
    };
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('blort');

    ws.getCell('A1').value = 'Hello, World!';
    ws.getCell('A2').value = 7;

    const buffer = await wb.xlsx.writeBuffer(options);

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(buffer.toString('base64'), options);
    const ws2 = wb2.getWorksheet('blort');

    expect(ws2.getCell('A1').value).to.equal('Hello, World!');
    expect(ws2.getCell('A2').value).to.equal(7);
  });
  it('should write csv via buffer', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('blort');

    ws.getCell('A1').value = 'Hello, World!';
    ws.getCell('B1').value = 'What time is it?';
    ws.getCell('A2').value = 7;
    ws.getCell('B2').value = '12pm';

    wb.csv.writeBuffer().then((buffer) => {
      expect(buffer.toString()).to.equal(
        '"Hello, World!",What time is it?\n7,12pm',
      );
    });
  });
});
