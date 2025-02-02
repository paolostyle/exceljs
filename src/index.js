import 'web-streams-polyfill/polyfill';
import * as enums from './doc/enums.ts';
import Workbook from './doc/workbook.js';
import WorkbookReader from './stream/xlsx/workbook-reader.js';
import WorkbookWriter from './stream/xlsx/workbook-writer.js';

const ExcelJS = {
  Workbook,
  stream: {
    xlsx: {
      WorkbookWriter,
      WorkbookReader,
    },
  },
  ...enums,
};

export { Workbook };
export * from './doc/enums.ts';

export default ExcelJS;
