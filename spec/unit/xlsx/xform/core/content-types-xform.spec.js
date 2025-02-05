import fs from 'node:fs';

import testXformHelper from '../test-xform-helper';

import ContentTypesXform from '#lib/xlsx/xform/core/content-types-xform.js';

const expectations = [
  {
    title: 'Three Sheets with shared strings',
    create() {
      return new ContentTypesXform();
    },
    preparedModel: {
      worksheets: [{ id: 1 }, { id: 2 }, { id: 3 }],
      media: [],
      drawings: [],
      sharedStrings: { count: 1 },
    },
    xml: fs
      .readFileSync(`${__dirname}/data/content-types.01.xml`)
      .toString()
      .replace(/\r\n/g, '\n'),
    tests: ['render'],
  },
  {
    title: 'Images with shared strings',
    create() {
      return new ContentTypesXform();
    },
    preparedModel: {
      worksheets: [{ id: 1 }, { id: 2 }],
      media: [
        { type: 'image', extension: 'png' },
        { type: 'image', extension: 'jpg' },
      ],
      drawings: [],
      sharedStrings: { count: 1 },
    },
    xml: fs
      .readFileSync(`${__dirname}/data/content-types.02.xml`)
      .toString()
      .replace(/\r\n/g, '\n'),
    tests: ['render'],
  },
  {
    title: 'Three Sheets without shared strings',
    create() {
      return new ContentTypesXform();
    },
    preparedModel: {
      worksheets: [{ id: 1 }, { id: 2 }, { id: 3 }],
      media: [],
      drawings: [],
    },
    xml: fs
      .readFileSync(`${__dirname}/data/content-types.03.xml`)
      .toString()
      .replace(/\r\n/g, '\n'),
    tests: ['render'],
  },
  {
    title: 'Images without shared strings',
    create() {
      return new ContentTypesXform();
    },
    preparedModel: {
      worksheets: [{ id: 1 }, { id: 2, useSharedStrings: false }],
      media: [
        { type: 'image', extension: 'png' },
        { type: 'image', extension: 'jpg' },
      ],
      drawings: [],
    },
    xml: fs
      .readFileSync(`${__dirname}/data/content-types.04.xml`)
      .toString()
      .replace(/\r\n/g, '\n'),
    tests: ['render'],
  },
];

describe('ContentTypesXform', () => {
  testXformHelper(expectations);
});
