import { AppModule } from './app-module.js';

/** Calc module boundary. Existing Calc behavior remains compatible with the Office state model. */
export class CalcModule extends AppModule {
  reset() { return { rows: 20, cols: 12, data: {}, styles: {}, selected: 'A1', sheetIndex: 0, sheets: [{ name: 'Hoja 1', rows: 20, cols: 12, data: {}, styles: {} }] }; }
  async load() { return this.storage.readJSON('calc/workbook.json', this.reset()); }
  async save(value, path = this.path || 'calc/Libro de cálculo.tetord') { this.path = path; return this.storage.writeJSON(path, value); }
  async openNativeRecord(record, value) { this.path = record.path; this.office.calcPath = record.path; this.office.calc = this.office.normalizeCalc(value); this.office.open('calc'); this.office.renderCalc(); return true; }
}
