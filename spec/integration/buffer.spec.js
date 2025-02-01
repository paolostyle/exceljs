const ExcelJS = require('#lib');

describe('ExcelJS', () => {
  it('should read and write xlsx via binary buffer', done => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('blort');

    ws.getCell('A1').value = 'Hello, World!';
    ws.getCell('A2').value = 7;

    wb.xlsx
      .writeBuffer()
      .then(buffer => {
        const wb2 = new ExcelJS.Workbook();
        return wb2.xlsx.load(buffer).then(() => {
          const ws2 = wb2.getWorksheet('blort');

          expect(ws2.getCell('A1').value).to.equal('Hello, World!');
          expect(ws2.getCell('A2').value).to.equal(7);
          done();
        });
      });
  });
  it('should read and write xlsx via base64 buffer', done => {
    const options = {
      base64: true,
    };
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('blort');

    ws.getCell('A1').value = 'Hello, World!';
    ws.getCell('A2').value = 7;

    wb.xlsx
      .writeBuffer(options)
      .then(buffer => {
        const wb2 = new ExcelJS.Workbook();
        return wb2.xlsx.load(buffer.toString('base64'), options).then(() => {
          const ws2 = wb2.getWorksheet('blort');

          expect(ws2.getCell('A1').value).to.equal('Hello, World!');
          expect(ws2.getCell('A2').value).to.equal(7);
          done();
        });
      });
  });
  it('should write csv via buffer', done => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('blort');

    ws.getCell('A1').value = 'Hello, World!';
    ws.getCell('B1').value = 'What time is it?';
    ws.getCell('A2').value = 7;
    ws.getCell('B2').value = '12pm';

    wb.csv
      .writeBuffer()
      .then(buffer => {
        expect(buffer.toString()).to.equal(
          '"Hello, World!",What time is it?\n7,12pm'
        );
        done();
      });
  });
});
