import _ from '../../../utils/under-dash.js';
import BaseXform from '../base-xform.js';

class SheetFormatPropertiesXform extends BaseXform {
  get tag() {
    return 'sheetFormatPr';
  }

  render(xmlStream, model) {
    if (model) {
      const attributes = {
        defaultRowHeight: model.defaultRowHeight,
        outlineLevelRow: model.outlineLevelRow,
        outlineLevelCol: model.outlineLevelCol,
        'x14ac:dyDescent': model.dyDescent,
      };
      if (model.defaultColWidth) {
        attributes.defaultColWidth = model.defaultColWidth;
      }

      // default value for 'defaultRowHeight' is 15, this should not be 'custom'
      if (!model.defaultRowHeight || model.defaultRowHeight !== 15) {
        attributes.customHeight = '1';
      }

      if (_.some(attributes, (value) => value !== undefined)) {
        xmlStream.leafNode('sheetFormatPr', attributes);
      }
    }
  }

  parseOpen(node) {
    if (node.name === 'sheetFormatPr') {
      this.model = {
        defaultRowHeight: Number.parseFloat(
          node.attributes.defaultRowHeight || '0',
        ),
        dyDescent: Number.parseFloat(node.attributes['x14ac:dyDescent'] || '0'),
        outlineLevelRow: Number.parseInt(
          node.attributes.outlineLevelRow || '0',
          10,
        ),
        outlineLevelCol: Number.parseInt(
          node.attributes.outlineLevelCol || '0',
          10,
        ),
      };
      if (node.attributes.defaultColWidth) {
        this.model.defaultColWidth = Number.parseFloat(
          node.attributes.defaultColWidth,
        );
      }
      return true;
    }
    return false;
  }

  parseText() {}

  parseClose() {
    return false;
  }
}

export default SheetFormatPropertiesXform;
