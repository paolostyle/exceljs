import { ZipReaderStream } from '@zip.js/zip.js';
import { loadFileToStream, writeFileFromStream } from '../../utils/files.js';

import { EventEmitter } from 'node:events';
import fs from 'node:fs';
import tmp from 'tmp';
import parseSax from '../../utils/parse-sax.js';
import WorkbookXform from '../../xlsx/xform/book/workbook-xform.js';
import RelationshipsXform from '../../xlsx/xform/core/relationships-xform.js';
import StyleManager from '../../xlsx/xform/style/styles-xform.js';
import HyperlinkReader from './hyperlink-reader.js';
import WorksheetReader from './worksheet-reader.js';

tmp.setGracefulCleanup();

class WorkbookReader extends EventEmitter {
  constructor(input, options = {}) {
    super();

    this.input = input;

    this.options = {
      worksheets: 'emit',
      sharedStrings: 'cache',
      hyperlinks: 'ignore',
      styles: 'ignore',
      entries: 'ignore',
      ...options,
    };

    this.styles = new StyleManager();
    this.styles.init();
  }

  async read(input, options) {
    try {
      for await (const { eventType, value } of this.parse(input, options)) {
        switch (eventType) {
          case 'shared-strings':
            this.emit(eventType, value);
            break;
          case 'worksheet':
            this.emit(eventType, value);
            await value.read();
            break;
          case 'hyperlinks':
            this.emit(eventType, value);
            break;
        }
      }
      this.emit('end');
      this.emit('finished');
    } catch (error) {
      this.emit('error', error);
    }
  }

  async *[Symbol.asyncIterator]() {
    for await (const { eventType, value } of this.parse()) {
      if (eventType === 'worksheet') {
        yield value;
      }
    }
  }

  async *parse(input, options) {
    if (options) this.options = options;

    this.stream = await loadFileToStream(input);
    // worksheets, deferred for parsing after shared strings reading
    const waitingWorkSheets = [];

    for await (const entry of this.stream.pipeThrough(new ZipReaderStream())) {
      switch (entry.filename) {
        case '_rels/.rels':
          break;
        case 'xl/_rels/workbook.xml.rels':
          await this._parseRels(entry.readable);
          break;
        case 'xl/workbook.xml':
          await this._parseWorkbook(entry.readable);
          break;
        case 'xl/sharedStrings.xml':
          yield* this._parseSharedStrings(entry.readable);
          break;
        case 'xl/styles.xml':
          await this._parseStyles(entry.readable);
          break;
        default: {
          const worksheetMatch = entry.filename.match(
            /xl\/worksheets\/sheet(\d+)[.]xml/,
          );
          if (worksheetMatch) {
            const sheetNo = worksheetMatch[1];
            if (this.sharedStrings && this.workbookRels) {
              yield* this._parseWorksheet(entry.readable, sheetNo);
            } else {
              // create temp file for each worksheet
              await new Promise((resolve, reject) => {
                tmp.file((err, path, fd, tempFileCleanupCallback) => {
                  if (err) {
                    return reject(err);
                  }
                  waitingWorkSheets.push({
                    sheetNo,
                    path,
                    tempFileCleanupCallback,
                  });

                  return writeFileFromStream(entry.readable, path)
                    .then(resolve)
                    .catch(reject);
                });
              });
            }
            break;
          }
          const hyperlinkMatch = entry.filename.match(
            /xl\/worksheets\/sheet(\d+)[.]xml/,
          );
          if (hyperlinkMatch) {
            yield* this._parseHyperlinks(entry.readable, hyperlinkMatch[1]);
          }
          break;
        }
      }
    }

    for (const {
      sheetNo,
      path,
      tempFileCleanupCallback,
    } of waitingWorkSheets) {
      const fileStream = fs.createReadStream(path);
      yield* this._parseWorksheet(fileStream, sheetNo);
      tempFileCleanupCallback();
    }
  }

  _emitEntry(payload) {
    if (this.options.entries === 'emit') {
      this.emit('entry', payload);
    }
  }

  async _parseRels(entry) {
    const xform = new RelationshipsXform();
    this.workbookRels = await xform.parseStream(entry);
  }

  async _parseWorkbook(entry) {
    if (!entry) return;

    this._emitEntry({ type: 'workbook' });

    const workbook = new WorkbookXform();
    await workbook.parseStream(entry);

    this.properties = workbook.map.workbookPr;
    this.model = workbook.model;
  }

  async *_parseSharedStrings(entry) {
    this._emitEntry({ type: 'shared-strings' });
    switch (this.options.sharedStrings) {
      case 'cache':
        this.sharedStrings = [];
        break;
      case 'emit':
        break;
      default:
        return;
    }

    let text = null;
    let richText = [];
    let index = 0;
    let font = null;
    for await (const events of parseSax(entry)) {
      for (const { eventType, value } of events) {
        if (eventType === 'opentag') {
          const node = value;
          switch (node.name) {
            case 'b':
              font = font || {};
              font.bold = true;
              break;
            case 'charset':
              font = font || {};
              font.charset = Number.parseInt(node.attributes.charset, 10);
              break;
            case 'color':
              font = font || {};
              font.color = {};
              if (node.attributes.rgb) {
                font.color.argb = node.attributes.argb;
              }
              if (node.attributes.val) {
                font.color.argb = node.attributes.val;
              }
              if (node.attributes.theme) {
                font.color.theme = node.attributes.theme;
              }
              break;
            case 'family':
              font = font || {};
              font.family = Number.parseInt(node.attributes.val, 10);
              break;
            case 'i':
              font = font || {};
              font.italic = true;
              break;
            case 'outline':
              font = font || {};
              font.outline = true;
              break;
            case 'rFont':
              font = font || {};
              font.name = node.value;
              break;
            case 'si':
              font = null;
              richText = [];
              text = null;
              break;
            case 'sz':
              font = font || {};
              font.size = Number.parseInt(node.attributes.val, 10);
              break;
            case 'strike':
              break;
            case 't':
              text = null;
              break;
            case 'u':
              font = font || {};
              font.underline = true;
              break;
            case 'vertAlign':
              font = font || {};
              font.vertAlign = node.attributes.val;
              break;
          }
        } else if (eventType === 'text') {
          text = text ? text + value : value;
        } else if (eventType === 'closetag') {
          const node = value;
          switch (node.name) {
            case 'r':
              richText.push({
                font,
                text,
              });

              font = null;
              text = null;
              break;
            case 'si':
              if (this.options.sharedStrings === 'cache') {
                this.sharedStrings.push(richText.length ? { richText } : text);
              } else if (this.options.sharedStrings === 'emit') {
                yield {
                  index: index++,
                  text: richText.length ? { richText } : text,
                };
              }

              richText = [];
              font = null;
              text = null;
              break;
          }
        }
      }
    }
  }

  async _parseStyles(entry) {
    this._emitEntry({ type: 'styles' });
    if (this.options.styles === 'cache') {
      this.styles = new StyleManager();
      await this.styles.parseStream(entry);
    }
  }

  *_parseWorksheet(iterator, sheetNo) {
    this._emitEntry({ type: 'worksheet', id: sheetNo });
    const worksheetReader = new WorksheetReader({
      workbook: this,
      id: sheetNo,
      iterator,
      options: this.options,
    });

    const matchingRel = (this.workbookRels || []).find(
      (rel) => rel.Target === `worksheets/sheet${sheetNo}.xml`,
    );
    const matchingSheet =
      matchingRel &&
      (this.model.sheets || []).find((sheet) => sheet.rId === matchingRel.Id);
    if (matchingSheet) {
      worksheetReader.id = matchingSheet.id;
      worksheetReader.name = matchingSheet.name;
      worksheetReader.state = matchingSheet.state;
    }
    if (this.options.worksheets === 'emit') {
      yield { eventType: 'worksheet', value: worksheetReader };
    }
  }

  *_parseHyperlinks(iterator, sheetNo) {
    this._emitEntry({ type: 'hyperlinks', id: sheetNo });
    const hyperlinksReader = new HyperlinkReader({
      workbook: this,
      id: sheetNo,
      iterator,
      options: this.options,
    });
    if (this.options.hyperlinks === 'emit') {
      yield { eventType: 'hyperlinks', value: hyperlinksReader };
    }
  }
  static Options = {
    worksheets: ['emit', 'ignore'],
    sharedStrings: ['cache', 'emit', 'ignore'],
    hyperlinks: ['cache', 'emit', 'ignore'],
    styles: ['cache', 'ignore'],
    entries: ['emit', 'ignore'],
  };
}

export default WorkbookReader;
