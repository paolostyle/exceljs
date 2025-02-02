import _ from '../../../utils/under-dash.js';
import BaseXform from '../base-xform.js';

function booleanToXml(model) {
  return model ? '1' : undefined;
}
function pageOrderToXml(model) {
  switch (model) {
    case 'overThenDown':
      return model;
    default:
      return undefined;
  }
}
function cellCommentsToXml(model) {
  switch (model) {
    case 'atEnd':
    case 'asDisplyed':
      return model;
    default:
      return undefined;
  }
}
function errorsToXml(model) {
  switch (model) {
    case 'dash':
    case 'blank':
    case 'NA':
      return model;
    default:
      return undefined;
  }
}
function pageSizeToModel(value) {
  return value !== undefined ? Number.parseInt(value, 10) : undefined;
}

class PageSetupXform extends BaseXform {
  get tag() {
    return 'pageSetup';
  }

  render(xmlStream, model) {
    if (model) {
      const attributes = {
        paperSize: model.paperSize,
        orientation: model.orientation,
        horizontalDpi: model.horizontalDpi,
        verticalDpi: model.verticalDpi,
        pageOrder: pageOrderToXml(model.pageOrder),
        blackAndWhite: booleanToXml(model.blackAndWhite),
        draft: booleanToXml(model.draft),
        cellComments: cellCommentsToXml(model.cellComments),
        errors: errorsToXml(model.errors),
        scale: model.scale,
        fitToWidth: model.fitToWidth,
        fitToHeight: model.fitToHeight,
        firstPageNumber: model.firstPageNumber,
        useFirstPageNumber: booleanToXml(model.firstPageNumber),
        usePrinterDefaults: booleanToXml(model.usePrinterDefaults),
        copies: model.copies,
      };
      if (_.some(attributes, (value) => value !== undefined)) {
        xmlStream.leafNode(this.tag, attributes);
      }
    }
  }

  parseOpen(node) {
    switch (node.name) {
      case this.tag:
        this.model = {
          paperSize: pageSizeToModel(node.attributes.paperSize),
          orientation: node.attributes.orientation || 'portrait',
          horizontalDpi: Number.parseInt(
            node.attributes.horizontalDpi || '4294967295',
            10,
          ),
          verticalDpi: Number.parseInt(
            node.attributes.verticalDpi || '4294967295',
            10,
          ),
          pageOrder: node.attributes.pageOrder || 'downThenOver',
          blackAndWhite: node.attributes.blackAndWhite === '1',
          draft: node.attributes.draft === '1',
          cellComments: node.attributes.cellComments || 'None',
          errors: node.attributes.errors || 'displayed',
          scale: Number.parseInt(node.attributes.scale || '100', 10),
          fitToWidth: Number.parseInt(node.attributes.fitToWidth || '1', 10),
          fitToHeight: Number.parseInt(node.attributes.fitToHeight || '1', 10),
          firstPageNumber: Number.parseInt(
            node.attributes.firstPageNumber || '1',
            10,
          ),
          useFirstPageNumber: node.attributes.useFirstPageNumber === '1',
          usePrinterDefaults: node.attributes.usePrinterDefaults === '1',
          copies: Number.parseInt(node.attributes.copies || '1', 10),
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
}

export default PageSetupXform;
