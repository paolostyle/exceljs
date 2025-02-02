import BaseXform from '../base-xform.js';
import StaticXform from '../static-xform.js';
import BlipFillXform from './blip-fill-xform.js';
import NvPicPrXform from './nv-pic-pr-xform.js';

class PicXform extends BaseXform {
  constructor() {
    super();

    this.map = {
      'xdr:nvPicPr': new NvPicPrXform(),
      'xdr:blipFill': new BlipFillXform(),
      'xdr:spPr': new StaticXform(PicXform.SpPrStructure),
    };
  }

  get tag() {
    return 'xdr:pic';
  }

  prepare(model, options) {
    model.index = options.index + 1;
  }

  render(xmlStream, model) {
    xmlStream.openNode(this.tag);

    this.map['xdr:nvPicPr'].render(xmlStream, model);
    this.map['xdr:blipFill'].render(xmlStream, model);
    this.map['xdr:spPr'].render(xmlStream, model);

    xmlStream.closeNode();
  }

  parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case this.tag:
        this.reset();
        break;
      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        break;
    }
    return true;
  }

  parseText() {}

  parseClose(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.mergeModel(this.parser.model);
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        return false;
      default:
        // not quite sure how we get here!
        return true;
    }
  }

  static SpPrStructure = {
    tag: 'xdr:spPr',
    c: [
      {
        tag: 'a:xfrm',
        c: [
          { tag: 'a:off', $: { x: '0', y: '0' } },
          { tag: 'a:ext', $: { cx: '0', cy: '0' } },
        ],
      },
      {
        tag: 'a:prstGeom',
        $: { prst: 'rect' },
        c: [{ tag: 'a:avLst' }],
      },
    ],
  };
}

export default PicXform;
