const fs = require('node:fs');
const path = require('node:path');

const HrStopwatch = require('./utils/hr-stopwatch');

const { Workbook } = require('#lib');

const filename = process.argv[2];

const wb = new Workbook();
const ws = wb.addWorksheet('blort');

ws.getCell('B2').value = 'Hello, World!';

const imageId = wb.addImage({
  filename: path.join(__dirname, 'data/image2.png'),
  extension: 'png',
});
const backgroundId = wb.addImage({
  buffer: fs.readFileSync(path.join(__dirname, 'data/bubbles.jpg')),
  extension: 'jpeg',
});
ws.addImage(imageId, {
  // tl: { col: 1, row: 1 },
  tl: 'B2',
  ext: { width: 100, height: 100 },
});

ws.addBackgroundImage(backgroundId);

const stopwatch = new HrStopwatch();
stopwatch.start();
wb.xlsx
  .writeFile(filename)
  .then(() => {
    const micros = stopwatch.microseconds;
    console.log('Done.');
    console.log('Time taken:', micros);
  })
  .catch((error) => {
    console.error(error.stack);
  });
