'use strict';
const assert = require('assert');
global.window = global;
require('../src/clParser.js');
const parser = global.CLParser;

const noseCone = [
  ['', '', '', '', ''],
  ['', 'Please nest and cut four (4X) Nose Cone Uprights', '', '', ''],
  ['', '', '', '', ''],
  ['', 'Nose Cone Upright', '', '', ''],
  ['', '', '', '', ''],
  ['', 'Part Number', 'Description', 'Qty', 'Material'],
  ['', '', '', '', ''],
  ['', '', '', '', ''],
  ['', 'HIN30W - 85 A', 'NOSE CONE MOUNTING LUGS', 4, '4mm Ramort 500'],
  ['', 'HIN30W - 91 A', 'PIPE REPLACEMENT GUSSET', 1, '4mm Ramort 500'],
  ['', 'HIN30W - 92 C', 'BULLBAR VERTICAL', 2, '8mm Mild Steel Plate'],
  ['', 'HIN30W - 94 B', 'PIPE REPLACEMENT GUSSET STIFFENER', 1, '4mm Ramort 500']
];
const parsed = parser.parseSheet(noseCone, { sheet: 'BAT - Nose cone' });
assert.strictEqual(parsed.length, 4, 'Should parse all four nose-cone parts.');
assert.deepStrictEqual(parsed.map(x => x.qty), [4, 1, 2, 1], 'Per-vehicle quantities must not be multiplied.');
assert.deepStrictEqual(parsed.map(x => x.batchMultiplier), [4, 4, 4, 4], 'Instruction-line 4X batch multiplier must apply to every row.');
assert.deepStrictEqual(parsed.map(x => x.effectiveQty), [16, 4, 8, 4], 'Batch total should be per-vehicle qty × batch multiplier.');
assert.strictEqual(parsed[0].description, 'NOSE CONE MOUNTING LUGS', 'Part description should be retained from the CL.');
assert.strictEqual(parsed[0].material, '4mm Ramor 500', 'Known CL material spelling normalization should remain active.');

const explicitColumn = parser.parseSheet([
  ['Part', 'Description', 'Qty/Vehicle', 'Material', 'Batch Multiplier'],
  ['P-1', 'BRACKET', 2, '4mm Armox', 3]
], { sheet: 'BAT parts' })[0];
assert.strictEqual(explicitColumn.batchMultiplier, 3, 'Row batch column should override the sheet default.');
assert.strictEqual(explicitColumn.effectiveQty, 6);

const tabMultiplier = parser.parseSheet([
  ['Part', 'Qty', 'Material'],
  ['P-2', 5, '4mm Armox']
], { sheet: 'BAT parts (x5)' })[0];
assert.strictEqual(tabMultiplier.batchMultiplier, 5, 'Tab-name xN syntax must remain supported.');
assert.strictEqual(tabMultiplier.effectiveQty, 25);

assert.throws(() => parser.parseSheet([
  ['Please nest x2'],
  ['Part', 'Qty', 'Material'],
  ['P-3', 1, '4mm Armox']
], { sheet: 'BAT parts (x3)' }), /Conflicting batch multipliers/, 'Conflicting instructions must stop unsafe import.');

const a = { ...parsed[0] };
const b = { ...parsed[0], sourceRow: 25, qty: 1, baseQty: 1, effectiveQty: 2, batchMultiplier: 2 };
const consolidated = parser.consolidate([a, b]);
assert.strictEqual(consolidated.length, 1);
assert.strictEqual(consolidated[0].clReviewReason, 'CL BATCH VARIANTS', 'A part with different per-line multipliers must be flagged.');
console.log('CL parser batch/data tests: PASS');
