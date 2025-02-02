import 'web-streams-polyfill/polyfill';

import enums from './doc/enums.js';
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

export default ExcelJS;
