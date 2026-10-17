/**
 * Weather Analysis inputs the user may edit in the UI.
 *
 * Everything else on the page is a formula (exactly as in the Excel workbook) and
 * recalculates automatically.
 *
 *  - E7 / E8      : Hot weather (>37°C) and Cold weather (<2°C) "direct" impact
 *  - B31..B57     : "Direct" rainfall impacts for Factors 1-5
 *  - C31..C57     : "Consequential" rainfall impacts. In the workbook these are fixed
 *                   constants (there is no formula behind them), so they are shown
 *                   read-only. Flip CONSEQUENTIAL_EDITABLE to true to let users edit them.
 *  - Total (D)    : always =SUM(Direct:Consequential) – never editable.
 */
export const CONSEQUENTIAL_EDITABLE = false;

const DIRECT_ROWS = [31, 32, 33, 37, 38, 39, 43, 44, 45, 49, 50, 51, 55, 56, 57];

export const EDITABLE_CELLS = new Set([
  'E7',
  'E8',
  ...DIRECT_ROWS.map((r) => `B${r}`),
  ...(CONSEQUENTIAL_EDITABLE ? DIRECT_ROWS.map((r) => `C${r}`) : []),
]);

export const STORAGE_KEYS = {
  dump: 'inclement-weather:bom-dump',
  overrides: 'inclement-weather:overrides',
};
