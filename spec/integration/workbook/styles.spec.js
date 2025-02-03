import fs from 'node:fs';
import path from 'node:path';
import stream from 'readable-stream';
import ExcelJS from '#lib';
import testUtils, { getTempFileName } from '../../utils/index';

// =============================================================================
// Sample Data
const richTextSample = fs.readFileSync(
  path.resolve(__dirname, '../data/rich-text-sample.txt'),
  'utf-8',
);
const richTextSampleA1 = require('../data/rich-text-sample-a1.json');

// =============================================================================
// Tests

describe('Workbook', () => {
  describe('Styles', () => {
    it('row styles and columns properly', async () => {
      const testFileName = getTempFileName();
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('blort');

      ws.columns = [
        { header: 'A1', width: 10 },
        {
          header: 'B1',
          width: 20,
          style: {
            font: testUtils.styles.fonts.comicSansUdB16,
            alignment: testUtils.styles.alignments[1].alignment,
          },
        },
        { header: 'C1', width: 30 },
      ];

      ws.getRow(2).font = testUtils.styles.fonts.broadwayRedOutline20;

      ws.getCell('A2').value = 'A2';
      ws.getCell('B2').value = 'B2';
      ws.getCell('C2').value = 'C2';
      ws.getCell('A3').value = 'A3';
      ws.getCell('B3').value = 'B3';
      ws.getCell('C3').value = 'C3';

      await wb.xlsx.writeFile(testFileName);
      const wb2 = new ExcelJS.Workbook();
      const wb2_1 = await wb2.xlsx.readFile(testFileName);
      const ws2 = wb2_1.getWorksheet('blort');

      ['A1', 'B1', 'C1', 'A2', 'B2', 'C2', 'A3', 'B3', 'C3'].forEach(
        (address) => {
          expect(ws2.getCell(address).value).to.equal(address);
        },
      );
      expect(ws2.getCell('B1').font).to.deep.equal(
        testUtils.styles.fonts.comicSansUdB16,
      );
      expect(ws2.getCell('B1').alignment).to.deep.equal(
        testUtils.styles.alignments[1].alignment,
      );
      expect(ws2.getCell('A2').font).to.deep.equal(
        testUtils.styles.fonts.broadwayRedOutline20,
      );
      expect(ws2.getCell('B2').font).to.deep.equal(
        testUtils.styles.fonts.broadwayRedOutline20,
      );
      expect(ws2.getCell('C2').font).to.deep.equal(
        testUtils.styles.fonts.broadwayRedOutline20,
      );
      expect(ws2.getCell('B3').font).to.deep.equal(
        testUtils.styles.fonts.comicSansUdB16,
      );
      expect(ws2.getCell('B3').alignment).to.deep.equal(
        testUtils.styles.alignments[1].alignment,
      );
      expect(ws2.getColumn(2).font).to.deep.equal(
        testUtils.styles.fonts.comicSansUdB16,
      );
      expect(ws2.getColumn(2).alignment).to.deep.equal(
        testUtils.styles.alignments[1].alignment,
      );
      expect(ws2.getRow(2).font).to.deep.equal(
        testUtils.styles.fonts.broadwayRedOutline20,
      );
    });

    it('in-cell formats properly in xlsx file', () => {
      // Stream from input string
      const testData = Buffer.from(richTextSample, 'base64');

      // Initiate the source
      const bufferStream = new stream.PassThrough();

      // Write your buffer
      bufferStream.write(testData);
      bufferStream.end();

      const wb = new ExcelJS.Workbook();
      return wb.xlsx.read(bufferStream).then(() => {
        const ws = wb.worksheets[0];
        expect(ws.getCell('A1').value).to.deep.equal(richTextSampleA1);
        expect(ws.getCell('A1').text).to.equal(ws.getCell('A2').value);
      });
    });

    it('null cells retain style', () => {
      const testFileName = getTempFileName();
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('blort');

      // one value here
      ws.getCell('B2').value = 'hello';

      // style here
      ws.getCell('B4').fill = testUtils.styles.fills.redDarkVertical;
      ws.getCell('B4').font = testUtils.styles.fonts.broadwayRedOutline20;

      return wb.xlsx
        .writeFile(testFileName)
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(testFileName);
        })
        .then((wb2) => {
          const ws2 = wb2.getWorksheet('blort');

          expect(ws2.getCell('B4').fill).to.deep.equal(
            testUtils.styles.fills.redDarkVertical,
          );
          expect(ws2.getCell('B4').font).to.deep.equal(
            testUtils.styles.fonts.broadwayRedOutline20,
          );
        });
    });

    it('sets row styles', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('basket');

      ws.getCell('A1').value = 5;
      ws.getCell('A1').numFmt = testUtils.styles.numFmts.numFmt1;
      ws.getCell('A1').font = testUtils.styles.fonts.arialBlackUI14;

      ws.getCell('C1').value = 'Hello, World!';
      ws.getCell('C1').alignment = testUtils.styles.namedAlignments.bottomRight;
      ws.getCell('C1').border = testUtils.styles.borders.doubleRed;
      ws.getCell('C1').fill = testUtils.styles.fills.redDarkVertical;

      ws.getRow(1).numFmt = testUtils.styles.numFmts.numFmt2;
      ws.getRow(1).font = testUtils.styles.fonts.comicSansUdB16;
      ws.getRow(1).alignment = testUtils.styles.namedAlignments.middleCentre;
      ws.getRow(1).border = testUtils.styles.borders.thin;
      ws.getRow(1).fill = testUtils.styles.fills.redGreenDarkTrellis;

      expect(ws.getCell('A1').numFmt).to.equal(
        testUtils.styles.numFmts.numFmt2,
      );
      expect(ws.getCell('A1').font).to.deep.equal(
        testUtils.styles.fonts.comicSansUdB16,
      );
      expect(ws.getCell('A1').alignment).to.deep.equal(
        testUtils.styles.namedAlignments.middleCentre,
      );
      expect(ws.getCell('A1').border).to.deep.equal(
        testUtils.styles.borders.thin,
      );
      expect(ws.getCell('A1').fill).to.deep.equal(
        testUtils.styles.fills.redGreenDarkTrellis,
      );

      expect(ws.findCell('B1')).toBeUndefined();

      expect(ws.getCell('C1').numFmt).to.equal(
        testUtils.styles.numFmts.numFmt2,
      );
      expect(ws.getCell('C1').font).to.deep.equal(
        testUtils.styles.fonts.comicSansUdB16,
      );
      expect(ws.getCell('C1').alignment).to.deep.equal(
        testUtils.styles.namedAlignments.middleCentre,
      );
      expect(ws.getCell('C1').border).to.deep.equal(
        testUtils.styles.borders.thin,
      );
      expect(ws.getCell('C1').fill).to.deep.equal(
        testUtils.styles.fills.redGreenDarkTrellis,
      );

      // when we 'get' the previously null cell, it should inherit the row styles
      expect(ws.getCell('B1').numFmt).to.equal(
        testUtils.styles.numFmts.numFmt2,
      );
      expect(ws.getCell('B1').font).to.deep.equal(
        testUtils.styles.fonts.comicSansUdB16,
      );
      expect(ws.getCell('B1').alignment).to.deep.equal(
        testUtils.styles.namedAlignments.middleCentre,
      );
      expect(ws.getCell('B1').border).to.deep.equal(
        testUtils.styles.borders.thin,
      );
      expect(ws.getCell('B1').fill).to.deep.equal(
        testUtils.styles.fills.redGreenDarkTrellis,
      );
    });

    it('sets col styles', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('basket');

      ws.getCell('A1').value = 5;
      ws.getCell('A1').numFmt = testUtils.styles.numFmts.numFmt1;
      ws.getCell('A1').font = testUtils.styles.fonts.arialBlackUI14;

      ws.getCell('A3').value = 'Hello, World!';
      ws.getCell('A3').alignment = testUtils.styles.namedAlignments.bottomRight;
      ws.getCell('A3').border = testUtils.styles.borders.doubleRed;
      ws.getCell('A3').fill = testUtils.styles.fills.redDarkVertical;

      ws.getColumn('A').numFmt = testUtils.styles.numFmts.numFmt2;
      ws.getColumn('A').font = testUtils.styles.fonts.comicSansUdB16;
      ws.getColumn('A').alignment =
        testUtils.styles.namedAlignments.middleCentre;
      ws.getColumn('A').border = testUtils.styles.borders.thin;
      ws.getColumn('A').fill = testUtils.styles.fills.redGreenDarkTrellis;

      expect(ws.getCell('A1').numFmt).to.equal(
        testUtils.styles.numFmts.numFmt2,
      );
      expect(ws.getCell('A1').font).to.deep.equal(
        testUtils.styles.fonts.comicSansUdB16,
      );
      expect(ws.getCell('A1').alignment).to.deep.equal(
        testUtils.styles.namedAlignments.middleCentre,
      );
      expect(ws.getCell('A1').border).to.deep.equal(
        testUtils.styles.borders.thin,
      );
      expect(ws.getCell('A1').fill).to.deep.equal(
        testUtils.styles.fills.redGreenDarkTrellis,
      );

      expect(ws.findRow(2)).toBeUndefined();

      expect(ws.getCell('A3').numFmt).to.equal(
        testUtils.styles.numFmts.numFmt2,
      );
      expect(ws.getCell('A3').font).to.deep.equal(
        testUtils.styles.fonts.comicSansUdB16,
      );
      expect(ws.getCell('A3').alignment).to.deep.equal(
        testUtils.styles.namedAlignments.middleCentre,
      );
      expect(ws.getCell('A3').border).to.deep.equal(
        testUtils.styles.borders.thin,
      );
      expect(ws.getCell('A3').fill).to.deep.equal(
        testUtils.styles.fills.redGreenDarkTrellis,
      );

      // when we 'get' the previously null cell, it should inherit the column styles
      expect(ws.getCell('A2').numFmt).to.equal(
        testUtils.styles.numFmts.numFmt2,
      );
      expect(ws.getCell('A2').font).to.deep.equal(
        testUtils.styles.fonts.comicSansUdB16,
      );
      expect(ws.getCell('A2').alignment).to.deep.equal(
        testUtils.styles.namedAlignments.middleCentre,
      );
      expect(ws.getCell('A2').border).to.deep.equal(
        testUtils.styles.borders.thin,
      );
      expect(ws.getCell('A2').fill).to.deep.equal(
        testUtils.styles.fills.redGreenDarkTrellis,
      );
    });
  });
});
