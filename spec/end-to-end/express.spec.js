import express from 'express';
import got from 'got';
import { PassThrough } from 'readable-stream';
import Excel from '#lib';
import testutils from '../utils/index';

describe('Express', () => {
  let server;
  beforeAll(() => {
    const app = express();
    app.get('/workbook', (req, res) => {
      const wb = testutils.createTestBook(new Excel.Workbook(), 'xlsx');
      res.setHeader(
        'Content-Type',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      );
      res.setHeader('Content-Disposition', 'attachment; filename=Report.xlsx');
      wb.xlsx.write(res).then(() => {
        res.end();
      });
    });
    server = app.listen(3003);
  });

  afterAll(() => {
    server.close();
  });

  it('downloads a workbook', async () => {
    const res = got.stream('http://127.0.0.1:3003/workbook', {
      decompress: false,
    });
    const wb2 = new Excel.Workbook();
    // TODO: Remove passThrough with got 10+ (requires node v10+)
    await wb2.xlsx.read(res.pipe(new PassThrough()));
    testutils.checkTestBook(wb2, 'xlsx');
  });
});
