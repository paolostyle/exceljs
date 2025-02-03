import tools from '../tools.js';

const self = {
  conditionalFormattings: tools.fix(
    require('../data/conditional-formatting.json'),
  ),
  getConditionalFormatting(type) {
    return self.conditionalFormattings[type] || null;
  },
  addSheet(wb) {
    const ws = wb.addWorksheet('conditional-formatting');
    const { types } = self.conditionalFormattings;
    types.forEach((type) => {
      const conditionalFormatting = self.getConditionalFormatting(type);
      if (conditionalFormatting) {
        ws.addConditionalFormatting(conditionalFormatting);
      }
    });
  },

  checkSheet(wb) {
    const ws = wb.getWorksheet('conditional-formatting');
    expect(ws).not.toBeUndefined();
    expect(ws.conditionalFormattings).not.toBeUndefined();
    (ws.conditionalFormattings && ws.conditionalFormattings).forEach((item) => {
      const type = item.rules?.[0].type;
      const conditionalFormatting = self.getConditionalFormatting(type);
      expect(item).to.have.property('ref');
      expect(item).to.have.property('rules');
      expect(self.conditionalFormattings[type]).to.have.property('ref');
      expect(self.conditionalFormattings[type]).to.have.property('rules');
      expect(item.ref).to.deep.equal(conditionalFormatting.ref);
      expect(item.rules.length).to.equal(conditionalFormatting.rules.length);
    });
  },
};

export default self;
