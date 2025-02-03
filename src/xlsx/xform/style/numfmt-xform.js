import _ from '../../../utils/under-dash.js';
import { defaultNumFormats } from '../../default-num-formats.ts';
import BaseXform from '../base-xform.js';

function hashDefaultFormats() {
  const hash = {};
  _.each(defaultNumFormats, (dnf, id) => {
    if (dnf.f) {
      hash[dnf.f] = Number.parseInt(id, 10);
    }
    // at some point, add the other cultures here...
  });
  return hash;
}
const defaultFmtHash = hashDefaultFormats();

// NumFmt encapsulates translation between number format and xlsx
class NumFmtXform extends BaseXform {
  constructor(id, formatCode) {
    super();

    this.id = id;
    this.formatCode = formatCode;
  }

  get tag() {
    return 'numFmt';
  }

  render(xmlStream, model) {
    xmlStream.leafNode('numFmt', {
      numFmtId: model.id,
      formatCode: model.formatCode,
    });
  }

  parseOpen(node) {
    switch (node.name) {
      case 'numFmt':
        this.model = {
          id: Number.parseInt(node.attributes.numFmtId, 10),
          formatCode: node.attributes.formatCode.replace(/[\\](.)/g, '$1'),
        };
        return true;
      default:
        return false;
    }
  }

  parseText() {}

  parseClose() {
    return false;
  }

  static getDefaultFmtId(formatCode) {
    return defaultFmtHash[formatCode];
  }

  static getDefaultFmtCode(numFmtId) {
    return defaultNumFormats[numFmtId]?.f;
  }
}

export default NumFmtXform;
