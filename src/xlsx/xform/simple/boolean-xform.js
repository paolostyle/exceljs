import BaseXform from '../base-xform.js';

class BooleanXform extends BaseXform {
  constructor(options) {
    super();

    this.tag = options.tag;
    this.attr = options.attr;
  }

  render(xmlStream, model) {
    if (model) {
      xmlStream.openNode(this.tag);
      xmlStream.closeNode();
    }
  }

  parseOpen(node) {
    if (node.name === this.tag) {
      this.model = true;
    }
  }

  parseText() {}

  parseClose() {
    return false;
  }
}

export default BooleanXform;
