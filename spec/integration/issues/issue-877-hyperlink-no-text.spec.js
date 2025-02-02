import fs from 'node:fs';
import ExcelJS from '#lib';
import { getTempFileName } from '../../utils';

describe('github issues', () => {
  it('issue 877 - hyperlink without text crashes on write', () => {
    const wb = new ExcelJS.Workbook();
    return wb.xlsx
      .readFile('./spec/integration/data/test-issue-877.xlsx')
      .then(() => {
        wb.xlsx
          .writeBuffer({
            useStyles: true,
            useSharedStrings: true,
          })
          .then((buffer) => {
            const wstream = fs.createWriteStream(getTempFileName());
            wstream.write(buffer);
            wstream.end();
          });
      });
  });
});
