import BaseXform from '../base-xform.js';

class WorkbookPivotCacheXform extends BaseXform {
  render(xmlStream, model) {
    xmlStream.leafNode('pivotCache', {
      cacheId: model.cacheId,
      'r:id': model.rId,
    });
  }

  parseOpen(node) {
    if (node.name === 'pivotCache') {
      this.model = {
        cacheId: node.attributes.cacheId,
        rId: node.attributes['r:id'],
      };
      return true;
    }
    return false;
  }

  parseText() {}

  parseClose() {
    return false;
  }
}

export default WorkbookPivotCacheXform;
