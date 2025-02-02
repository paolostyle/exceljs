import { each } from '#lib/utils/under-dash.js';

export class ColumnSum {
  constructor(columns) {
    this.columns = columns;
    this.sums = [];
    this.count = 0;
    each(this.columns, (column) => {
      this.sums[column] = 0;
    });
  }

  add(row) {
    each(this.columns, (column) => {
      this.sums[column] += row.getCell(column).value;
    });
    this.count++;
  }

  toString() {
    return this.sums.join(', ');
  }

  toAverages() {
    return this.sums
      .map((value) => (value ? value / this.count : value))
      .join(', ');
  }
}
