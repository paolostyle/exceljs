const fs = require('node:fs/promises');
const JSZip = require('jszip');

const filename = process.argv[2];

const jsZip = new JSZip();

fs.readFileAsync(filename)
  .then((data) => {
    console.log('data', data);
    return jsZip.loadAsync(data);
  })
  .then((zip) => {
    zip.forEach((path, entry) => {
      if (!entry.dir) {
        // console.log(path, entry)
        console.log(
          path,
          entry.name,
          entry._data.compressedSize,
          entry._data.uncompressedSize,
        );
      }
    });
  });
