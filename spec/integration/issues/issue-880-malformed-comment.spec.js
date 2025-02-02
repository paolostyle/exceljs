import fs from 'node:fs';

import ExcelJS from '#lib';

// this file to contain integration tests created from github issues
const TEST_XLSX_FILE_NAME = './spec/out/wb-issue-880.test.xlsx';

describe('github issues', () => {
  it(
    'issue 880 - malformed comment crashes on write',
    { timeout: 6000 },
    async () => {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.readFile('./spec/integration/data/test-issue-880.xlsx');
      wb.xlsx
        .writeBuffer({
          useStyles: true,
          useSharedStrings: true,
        })
        .then((buffer) => {
          const wstream = fs.createWriteStream(TEST_XLSX_FILE_NAME);
          wstream.write(buffer);
          wstream.end();
        });
    },
  );
});
