const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const source = fs.readFileSync(__dirname + '/app.js', 'utf8');
const values = {};
const output = {};
const elements = {};
let rows = [];
const jq = (id) => ({
  val(v) { if (v !== undefined) values[id] = v; return values[id]; },
  text(v) { output[id] = v; }, html() {return this;}, append() {}, show() {}, hide() {}, toggleClass() {}, toggle() {}
});
const context = vm.createContext({console, $, window: {}, document: {
  getElementById(id) { return elements[id] ||= {value: '', setCustomValidity(v) {this.error = v;}, reportValidity() {}}; },
  querySelectorAll() {return rows;}, createElement() {return {dataset:{}};}
}, Swal: {fire() {}}});
function $(id) { return jq(id); }
vm.runInContext(source.slice(source.indexOf('// ===== Shared Utilities'), source.indexOf('// FILE IMPORT ENGINE')), context);
vm.runInContext(source.slice(source.indexOf('const COST ='), source.indexOf('let costHistory')), context);
vm.runInContext(source.slice(source.indexOf('function getCostCapacity'), source.indexOf('function createBookTable')), context);
vm.runInContext(source.slice(source.indexOf('function digitalPaperGroups'), source.indexOf('function buildDigitalPrintPreview')), context);
vm.runInContext(source.slice(source.indexOf('function calculateCostBooks'), source.indexOf('function saveCostHistory')), context);
vm.runInContext('let prodHolidays = [];', context);
vm.runInContext(source.slice(source.indexOf('function calcFinishDate'), source.indexOf('// Function to update speed badge')), context);
for (const [n, expected] of [[0,0],[0.01,0.5],[0.5,0.5],[1,1],[1.1,1.5],[1.23,1.5],[1.5,1.5],[2.1,2.5],[2.5,2.5]]) {
  assert.equal(vm.runInContext(`roundUpHalfDay(${n})`, context), expected);
}
Object.assign(values, {'#numBooks':1, '#costTitle1':'Book', '#costPages1':22, '#costQty1':100, '#costMaterial1':'WF', '#costDailyCapacity':1000});
for (const size of ['A4','A3']) {
  values['#costPaper1'] = size;
  vm.runInContext('calculateCostBooks()', context);
  assert.equal(output['#costTotalDays'], 1.5, size + ' physical sheets');
}
const extra = elements.prodExtraDays = {value:'0.5', setCustomValidity(v) {this.error=v;}, reportValidity() {}};
elements.prodStartDate = {value:'2026-10-13'};
elements.prodFormsPerDay = {value:'10'};
elements.prodDetailTable = {innerHTML:'', appendChild() {}, querySelectorAll() {return [];}};
const makeRow = () => ({classList:{remove(){}, add(){}}, querySelector(s) {return {value: ({'.prod-title':'Offset', '.prod-pages':'88', '.prod-qty':'100000', '.prod-material':'WF', '.prod-binding':'Perfect'})[s]};}});
rows = [makeRow()];
vm.runInContext('runProductionCalc()', context);
assert.equal(context.window._prodData.totalDaysAll, 2, '1.5 printing + 0.5 extra, independent of sheets');
rows = [makeRow(), makeRow()];
vm.runInContext('runProductionCalc()', context);
assert.equal(context.window._prodData.totalDaysAll, 3, 'round aggregate 22/10 to 2.5, then add 0.5');
for (const invalid of ['1.1','1.23','-0.5','Infinity']) {
  extra.value = invalid;
  vm.runInContext('runProductionCalc()', context);
  assert.ok(extra.error);
  assert.equal(context.window._prodData, null);
}
extra.value = '2.5';
vm.runInContext('runProductionCalc()', context);
assert.equal(extra.error, '');
assert.equal(context.window._prodData.totalDaysAll, 5);
console.log('PASS: half-day rounding, A4/A3 physical sheets, offset child capacity, aggregate totals, extra-day validation.');
