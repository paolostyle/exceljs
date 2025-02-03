import _ from './under-dash.js';

const tools = {
  dtMatcher: /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}[.]\d{3}Z$/,
  fix: function fix(o) {
    // clone the object and replace any date-like strings with new Date()
    let clone;
    if (Array.isArray(o)) {
      clone = [];
    } else if (typeof o === 'object') {
      clone = {};
    } else if (typeof o === 'string' && tools.dtMatcher.test(o)) {
      return new Date(o);
    } else {
      return o;
    }
    _.each(o, (value, name) => {
      if (value !== undefined) {
        clone[name] = fix(value);
      }
    });
    return clone;
  },

  concatenateFormula(...args) {
    const values = args.map((value) => `"${value}"`);
    return {
      formula: `CONCATENATE(${values.join(',')})`,
    };
  },
};

export default tools;
