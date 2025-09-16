/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office, localStorage, setTimeout */

// Globální proměnné pro konfiguraci vzorkovače
let reliabilityConfig = {
  confidence: 95,
  expectedError: 5,
  populationSize: 1000
};
// Flag whether the reliability factor has been explicitly set by the user
let reliabilitySet = false;

// Remember last computed total (SUM of ABS) for Task 5
let totalSumABS = null;
// Requested/confirmed sample size (user-chosen)
let requestedSampleSize = null;
let requestedSampleReason = null;
// Seed for reproducible RNG (Task6)
let task6Seed = null;

// Import FS_PARAMETRY (in-code mapping of combinations to numeric factor)
import fsParams from './FS_Parametry';

// Optimalized batch size for large datasets
const BATCH_SIZE = 1500; // Further reduced for Excel API limits
const LARGE_DATASET_THRESHOLD = 50000; // Warn user for datasets larger than this
const MEMORY_CRITICAL_THRESHOLD = 200000; // Extra precautions for very large datasets
const OUTPUT_BATCH_SIZE = 800; // Smaller batches for output operations

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    
  // Nastavení event listenerů (s kontrolou existence prvků)
  const setReliabilityBtn = document.getElementById("set-reliability");
  if (setReliabilityBtn) setReliabilityBtn.onclick = openReliabilityDialog;

  // wire reset sampler button
  const resetBtnMain = document.getElementById('reset-sampler');
  if (resetBtnMain) resetBtnMain.onclick = resetSampler;

  const selectRangeBtn = document.getElementById("select-range");
  if (selectRangeBtn) selectRangeBtn.onclick = selectDataRange;

  // New button to select column for value column (task 4)
  const selectColumnBtn = document.getElementById("select-column");
  if (selectColumnBtn) selectColumnBtn.onclick = selectValueColumn;

  // Task 5: compute sum button wiring
  const task5Btn = document.getElementById('task5-compute-sum');
  if (task5Btn) task5Btn.onclick = computeTask5Sum;

  // Task 5: compute sample size wiring
  const task5SizeBtn = document.getElementById('task5-calc-size');
  if (task5SizeBtn) task5SizeBtn.onclick = computeTask5Size;

  // Task 5: choice whether to use calculated size
  const task5UseSelect = document.getElementById('task5-use-calculated');
  if (task5UseSelect) task5UseSelect.onchange = handleTask5UseChoice;
  // confirm button
  const task5ConfirmBtn = document.getElementById('task5-confirm-size');
  if (task5ConfirmBtn) task5ConfirmBtn.onclick = confirmRequestedSize;
  // wire sampling method change to update Task6 method display
  const samplingMethodEl = document.getElementById('sampling-method');
  if (samplingMethodEl) samplingMethodEl.onchange = () => {
    const m = document.getElementById('task6-method-display');
    if (m) m.value = samplingMethodEl.options[samplingMethodEl.selectedIndex].text;
  };
  // also populate on load with current selection
  try {
    const m0 = document.getElementById('task6-method-display');
    const samplingMethodEl0 = document.getElementById('sampling-method');
    if (m0 && samplingMethodEl0) m0.value = samplingMethodEl0.options[samplingMethodEl0.selectedIndex].text;
  } catch (e) {}

  // Wire Task6: exclude-over-significance select
  const task6ExcludeSel = document.getElementById('task6-exclude-over-significance');
  if (task6ExcludeSel) {
    task6ExcludeSel.onchange = (ev) => {
      try { saveTask6ExcludeSetting(task6ExcludeSel.value); } catch (e) { console.warn('Nelze uložit volbu v Task6:', e); }
    };
  }
  // load persisted value
  try { loadTask6ExcludeSetting(); } catch (e) { /* ignore */ }
  try { loadTask6PrintTarget(); } catch (e) { /* ignore */ }
  // wire compute adjusted sum button if present
  const task6ComputeBtn = document.getElementById('task6-compute-adjusted-sum');
  if (task6ComputeBtn) task6ComputeBtn.onclick = computeTask6AdjustedSum;
  // wire print target select
  const printTargetSel = document.getElementById('task6-print-target');
  if (printTargetSel) {
    printTargetSel.onchange = () => { try { saveTask6PrintTarget(printTargetSel.value); } catch (e) { console.warn(e); } };
  }
  // wire generate sample button in Task6
  const task6GenBtn = document.getElementById('task6-generate-sample');
  if (task6GenBtn) task6GenBtn.onclick = task6GenerateSampleHandler;
  const task6ParamsBtn = document.getElementById('task6-generate-parameters');
  if (task6ParamsBtn) task6ParamsBtn.onclick = task6GenerateParametersHandler;

  // wire seed input
  const seedEl = document.getElementById('task6-seed');
  if (seedEl) seedEl.onchange = () => { try { saveTask6Seed(seedEl.value); } catch (e) { console.warn(e); } };
  try { loadTask6Seed(); } catch (e) {}

  // wire seed buttons
  const genSeedBtn = document.getElementById('task6-generate-seed');
  const regenSeedBtn = document.getElementById('task6-regenerate-seed');
  if (genSeedBtn) genSeedBtn.onclick = () => { const s = makeRandomSeed(); try { document.getElementById('task6-seed').value = s; saveTask6Seed(s); showNotification('Seed vygenerován.'); } catch (e) { console.warn(e); } };
  if (regenSeedBtn) regenSeedBtn.onclick = () => { const s = makeRandomSeed(); try { document.getElementById('task6-seed').value = s; saveTask6Seed(s); showNotification('Seed znovu vygenerován.'); } catch (e) { console.warn(e); } };

  // Significance (task 5) wiring
  const fillSignificanceBtn = document.getElementById('fill-significance');
  if (fillSignificanceBtn) fillSignificanceBtn.onclick = openSignificanceDialog;

  // Modal buttons for significance
  const sigSave = document.getElementById('significance-save-btn');
  const sigCancel = document.getElementById('significance-cancel-btn');
  const sigReset = document.getElementById('significance-reset-btn');
  if (sigSave) sigSave.onclick = handleSignificanceSave;
  if (sigCancel) sigCancel.onclick = () => { const m = document.getElementById('significance-modal'); if (m) m.style.display = 'none'; };
  if (sigReset) sigReset.onclick = handleSignificanceReset;
    
    // Inicializace s výchozími hodnotami
    // load persisted settings if any
    loadReliabilityFromStorage();
    updateReliabilityDisplay();

    // Load significance value from storage
    loadSignificanceFromStorage();
  // Load last computed total for Task 5 if present
  try { loadTotalSumAbsFromStorage(); } catch (e) {}

    // Wire modal buttons if present
    const modal = document.getElementById("reliability-modal");
    const saveBtn = document.getElementById("modal-save-btn");
    const cancelBtn = document.getElementById("modal-cancel-btn");
    const resetBtn = document.getElementById("modal-reset-btn");
    if (saveBtn) saveBtn.onclick = handleModalSave;
    if (cancelBtn) cancelBtn.onclick = () => {
      if (modal) modal.style.display = "none";
    };
    if (resetBtn) resetBtn.onclick = handleModalReset;
  }
});

// Funkce pro otevření dialogu faktoru spolehlivosti
// Open modal-based reliability dialog
async function openReliabilityDialog() {
  const modal = document.getElementById("reliability-modal");
  if (!modal) {
    showNotification("Dialog není dostupný.");
    return;
  }
  // populate selects with current values (if present)
  const controlEl = document.getElementById("reliability-control-risk");
  const inherentEl = document.getElementById("reliability-inherent-risk");
  const rmmEl = document.getElementById("reliability-rmm-level");
  const analyticalEl = document.getElementById("reliability-analytical-tests");
  const controlTestsEl = document.getElementById("reliability-control-tests");

  if (controlEl && reliabilityConfig.controlRisk) controlEl.value = reliabilityConfig.controlRisk;
  // Always set selects to stored value or empty string so placeholder '<Vyberte>' appears
  if (controlEl) controlEl.value = reliabilityConfig.controlRisk || '';
  if (inherentEl) inherentEl.value = reliabilityConfig.inherentRisk || '';
  if (rmmEl) rmmEl.value = reliabilityConfig.rmmLevel || '';
  if (analyticalEl) analyticalEl.value = reliabilityConfig.analyticalTests || '';
  if (controlTestsEl) controlTestsEl.value = reliabilityConfig.controlTests || '';

  modal.style.display = "flex";
}

// Reset all UI fields and stored settings to initial state
function resetSampler() {
  try {
    // Clear various UI inputs
    const idsToClear = [
  'reliability-factor', 'significance-type-display', 'significance-value-display', 'significance-justification-display',
  'sampling-method', 'data-range', 'value-column', 'task5-result-display', 'task5-size-display', 'task5-use-calculated', 'task5-override-size', 'task5-override-reason', 'task5-confirm-info',
  'task6-method-display', 'task6-final-size-display', 'task6-exclude-over-significance', 'task6-param-sum-display', 'task6-significant-count-display', 'task6-print-target', 'task6-seed'
    ];
    idsToClear.forEach(id => {
      try {
        const el = document.getElementById(id);
        if (!el) return;
        if (el.tagName === 'SELECT') { el.value = ''; }
        else el.value = '';
      } catch (e) {}
    });

    // hide optional areas
    const overrideArea = document.getElementById('task5-override-area'); if (overrideArea) overrideArea.style.display = 'none';
    const countArea = document.getElementById('task6-count-area'); if (countArea) countArea.style.display = 'none';
    const resultsDiv = document.getElementById('results'); if (resultsDiv) resultsDiv.style.display = 'none';

    // Clear stored configs
  const keys = ['reliabilityConfig','significanceConfig','TotalSumABS','task6ExcludeOverSignificance','task6PrintTarget','task6Seed'];
    keys.forEach(k => { try { localStorage.removeItem(k); } catch (e) {} });

    // Reset in-memory vars
    reliabilityConfig = { confidence:95, expectedError:5, populationSize:1000 };
    reliabilitySet = false;
    totalSumABS = null;
    requestedSampleSize = null;
    requestedSampleReason = null;

    // Update UI displays that need special handling
    // ensure sampling method select and Task6 method display show placeholder
    try {
      const samplingMethodEl = document.getElementById('sampling-method');
      if (samplingMethodEl) samplingMethodEl.value = '';
      const task6MethodEl = document.getElementById('task6-method-display');
      if (task6MethodEl) { task6MethodEl.value = '<Vyberte>'; }
    } catch (e) {}
    updateReliabilityDisplay();
    showNotification('Vzorkovač byl restartován.');
  } catch (e) {
    console.error('Chyba při resetu vzorkovače:', e);
    showNotification('Chyba při restartu vzorkovače.');
  }
}

// Default handler for Task6 "Vygenerovat výběr vzorku" button.
// Currently delegates to existing generateSample() function; can be extended per your design.
async function task6GenerateSampleHandler() {
  let success = false;
  try {
    const methodEl = document.getElementById('sampling-method');
    const method = methodEl ? methodEl.value : null;
    if (method === 'monetary-random-walk') {
      success = await generateMonetaryRandomWalkTask6();
      } else if (method === 'random-number-generator') {
        success = await generateRandomNumberGeneratorTask6();
      } else {
        await generateSample();
        success = true;
    }
  } catch (e) {
    console.error('Chyba při spuštění generování vzorku z Task6:', e);
    showNotification('Chyba při generování vzorku.');
  }
}

// Handler for the "Vygenerovat parametry vzorkovače" button
async function task6GenerateParametersHandler() {
  try {
    const target = _get_selected_value('task6-print-target');
    if (target && target !== '') {
      const params = gatherSamplerParameters();
      await printParameters(params, target);
      showNotification('Parametry byly vygenerovány.');
    } else {
      showNotification('Vyberte prosím cíl pro tisk parametrů.');
    }
  } catch (e) {
    console.error('Chyba při generování parametrů:', e);
    showNotification('Chyba při generování parametrů.');
  }
}

// Implementation for "Náhodný generátor čísel" invoked from Task6
async function generateRandomNumberGeneratorTask6() {
  try {
    const dataRange = document.getElementById('data-range') && document.getElementById('data-range').value;
    const valueColumn = document.getElementById('value-column') && document.getElementById('value-column').value;
    if (!dataRange || !valueColumn) { showNotification('Nejdříve vyberte oblast dat a sloupec (úloha 4).'); return false; }

    // exclude setting and significance
    const excludeSel = document.getElementById('task6-exclude-over-significance');
    const exclude = excludeSel && excludeSel.value === 'yes';
    let significance = null;
    try {
      const raw = localStorage.getItem('significanceConfig');
      if (raw) {
        const cfg = JSON.parse(raw);
        if (cfg && cfg.rawValue !== undefined && cfg.rawValue !== null) significance = Number(cfg.rawValue);
      }
    } catch (e) {}
    if (exclude && (significance === null || Number.isNaN(significance))) {
      showNotification('Není nastavena významnost, přepněte nebo uložte hladinu významnosti (úloha 3).');
      return false;
    }

    // determine sample size (prefer requestedSampleSize)
    let sampleSize = requestedSampleSize || null;
    if (!sampleSize) {
      const display = document.getElementById('task5-size-display');
      if (display && display.value) {
        const parsed = parseNumberLoose(display.value);
        if (parsed !== null) sampleSize = Math.ceil(parsed);
      }
    }
    if (!sampleSize || sampleSize <= 0) { showNotification('Neexistuje validní velikost vzorku. Nejprve ji potvrďte v úloze 5.'); return false; }

    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      // sanitize address
      let addr = dataRange;
      if (addr.indexOf('!') !== -1) addr = addr.split('!').pop();
      addr = addr.replace(/\$/g, '').trim();
      const range = worksheet.getRange(addr);
      range.load(['values','rowIndex','columnIndex','rowCount','columnCount']);
      await context.sync();

      const values = range.values;
      const absColIndex = getColumnIndex(String(valueColumn).toUpperCase());
      const rangeStartCol = typeof range.columnIndex === 'number' ? range.columnIndex : 0;
      const relColIndex = absColIndex - rangeStartCol;
      const rowCount = range.rowCount || (values ? values.length : 0);

      if (!Array.isArray(values) || values.length <= 1 || relColIndex < 0 || relColIndex >= (values[0] ? values[0].length : 0)) {
        showNotification('Zadaný sloupec není v rámci vybraného rozsahu dat.');
        return false;
      }

      const dataRows = Math.max(0, rowCount - 1); // exclude header
      if (dataRows === 0) { showNotification('Datová oblast neobsahuje žádné řádky.'); return false; }

      // step based on number of rows (not sum)
      const step = dataRows / sampleSize;
      if (!isFinite(step) || step <= 0) { showNotification('Neplatný krok (krok <= 0).'); return false; }

  // random start in [0, step)
  const rng = getTask6Rng();
  const start = (typeof rng === 'function' ? rng() : Math.random()) * step;

      // selection counts per data row (0-based index for data rows)
      const selCounts = new Array(dataRows).fill(0);

      for (let k = 0; k < sampleSize; k++) {
        const acc = start + k * step;
        // wrap-around using modulo of dataRows
        let posFloat = acc % dataRows;
        if (posFloat < 0) posFloat += dataRows;
        const posIndex = Math.floor(posFloat); // 0-based among data rows
        selCounts[posIndex] = (selCounts[posIndex] || 0) + 1;
      }

      // Write header row
      const newCols = 2;
      const writeColIndex = range.columnIndex + range.columnCount;
      const headerRange = worksheet.getRangeByIndexes(range.rowIndex, writeColIndex, 1, newCols);
      headerRange.values = [['Pořadí', 'Výběr']];
      headerRange.format.font.bold = true;

      // Process data rows in batches
      const batchGen = readRangeInBatches(context, worksheet, range, BATCH_SIZE);
      for await (const { values: batchValues, startRow: batchStartRow } of batchGen) {
        const isFirstBatch = batchStartRow === range.rowIndex;
        const startIdx = isFirstBatch ? 1 : 0; // skip header row
        const outputRows = [];
        for (let i = startIdx; i < batchValues.length; i++) {
          const absoluteRowIdx = (batchStartRow - range.rowIndex) + i; // 0‑based index within the range
          const order = absoluteRowIdx; // same as row number (header is 0)
          const absRaw = batchValues[i][relColIndex];
          const parsed = parseNumberLoose(absRaw);
          const absVal = parsed === null ? 0 : Math.abs(parsed);
          const dataRowZeroBased = absoluteRowIdx - 1;
          let selection = 'Ne';
          if (exclude && significance !== null && absVal > significance) {
            selection = 'Ano - významnost';
          } else if (selCounts[dataRowZeroBased] && selCounts[dataRowZeroBased] > 0) {
            const cnt = selCounts[dataRowZeroBased];
            selection = cnt === 1 ? 'Ano - NGČ' : `Ano - NGČ (x${cnt})`;
          }
          outputRows.push([order, selection]);
        }
        const writeStartRow = range.rowIndex + (isFirstBatch ? 1 : (batchStartRow - range.rowIndex));
        await writeRangeInBatches(context, worksheet, writeStartRow, writeColIndex, outputRows);
      }

      await context.sync();

  const usedSeed2 = task6Seed || (document.getElementById('task6-seed') && document.getElementById('task6-seed').value) || '';
  showNotification('Generování (NGČ) dokončeno. Výsledky jsou vloženy do listu.' + (usedSeed2 ? ` Použit seed: ${usedSeed2}` : ''));
    });
    return true;
  } catch (e) {
    console.error('Chyba v generateRandomNumberGeneratorTask6:', e);
    showNotification('Chyba při generování NGČ: ' + (e && e.message ? e.message : ''));
    return false;
  }
}

// New implementation for monetary random walk triggered from Task6
async function generateMonetaryRandomWalkTask6() {
  try {
    const dataRange = document.getElementById('data-range') && document.getElementById('data-range').value;
    const valueColumn = document.getElementById('value-column') && document.getElementById('value-column').value;
    if (!dataRange || !valueColumn) { showNotification('Nejdříve vyberte oblast dat a sloupec (úloha 4).'); return false; }

    // determine exclude setting and significance
    const excludeSel = document.getElementById('task6-exclude-over-significance');
    const exclude = excludeSel && excludeSel.value === 'yes';
    let significance = null;
    try {
      const raw = localStorage.getItem('significanceConfig');
      if (raw) {
        const cfg = JSON.parse(raw);
        if (cfg && cfg.rawValue !== undefined && cfg.rawValue !== null) significance = Number(cfg.rawValue);
      }
    } catch (e) {}
    if (exclude && (significance === null || Number.isNaN(significance))) {
      showNotification('Není nastavena významnost, přepněte nebo uložte hladinu významnosti (úloha 3).');
      return false;
    }

    // determine sample size (prefer requestedSampleSize)
    let sampleSize = requestedSampleSize || null;
    if (!sampleSize) {
      const display = document.getElementById('task5-size-display');
      if (display && display.value) {
        const parsed = parseNumberLoose(display.value);
        if (parsed !== null) sampleSize = Math.ceil(parsed);
      }
    }
    if (!sampleSize || sampleSize <= 0) { showNotification('Neexistuje validní velikost vzorku. Nejprve ji potvrďte v úloze 5.'); return false; }

    // read adjusted total sum from Task6 param display (or fallback to totalSumABS)
    let totalSum = null;
    try {
      const p = document.getElementById('task6-param-sum-display');
      if (p && p.value) {
        const parsed = parseNumberLoose(p.value);
        if (parsed !== null) totalSum = parsed;
      }
    } catch (e) {}
    if (totalSum === null) totalSum = totalSumABS;

    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      // sanitize address
      let addr = dataRange;
      if (addr.indexOf('!') !== -1) addr = addr.split('!').pop();
      addr = addr.replace(/\$/g, '').trim();
      const range = worksheet.getRange(addr);
      range.load(['values','rowIndex','columnIndex','rowCount','columnCount']);
      await context.sync();

      const values = range.values;
      const absColIndex = getColumnIndex(String(valueColumn).toUpperCase());
      const rangeStartCol = typeof range.columnIndex === 'number' ? range.columnIndex : 0;
      const relColIndex = absColIndex - rangeStartCol;
      const rowCount = range.rowCount || (values ? values.length : 0);

      if (!Array.isArray(values) || values.length <= 1 || relColIndex < 0 || relColIndex >= (values[0] ? values[0].length : 0)) {
        showNotification('Zadaný sloupec není v rámci vybraného rozsahu dat.');
        return false;
      }

      // fallback compute totalSum if missing
      if (totalSum === null || Number.isNaN(totalSum)) {
        totalSum = getTotalAbsoluteSum(values, relColIndex);
      }
      if (!totalSum || Number.isNaN(totalSum) || totalSum <= 0) { showNotification('TotalSum je nula nebo neplatná.'); return false; }

  const step = totalSum / sampleSize;
  if (!isFinite(step) || step <= 0) { showNotification('Neplatný krok (krok <= 0).'); return false; }
  const rng = getTask6Rng();
  const start = (typeof rng === 'function' ? rng() : Math.random()) * step; // real random start in [0,step)

      // prepare targets for NPP picks (sampleSize targets)
      const targets = new Array(sampleSize);
      for (let j = 0; j < sampleSize; j++) targets[j] = start + j * step;
      let targetIndex = 0;

      // prepare output array (rowCount rows x 5 cols)
      const newCols = 5;
      const out = new Array(rowCount);
      // header row
      out[0] = ['ABS', 'Kumulace ABS', 'Index/krok', 'Výběr', 'Důvod'];

      // First pass: compute abs values and mark significance selections (exclude logic)
      let significantCount = 0;
      let nppCount = 0;
      const absVals = new Array(rowCount);
      const addVals = new Array(rowCount);
      const selections = new Array(rowCount);
      const reasons = new Array(rowCount);

      for (let i = 0; i < rowCount; i++) {
        selections[i] = 'Ne';
        reasons[i] = '';
      }

      for (let i = 1; i < rowCount; i++) {
        const raw = values[i][relColIndex];
        const parsed = parseNumberLoose(raw);
        const absVal = parsed === null ? 0 : Math.abs(parsed);
        absVals[i] = absVal;
        // default addVal is the abs value
        let addVal = absVal;
        // if exclude and exceeds significance, mark as significance and set addVal to 0
        if (exclude && significance !== null && absVal > significance) {
          addVal = 0;
          selections[i] = 'Ano - významnost';
          reasons[i] = 'přesah významnosti';
          significantCount++;
        }
        addVals[i] = addVal;
      }

      // Second pass: compute cumulative and assign NPP based on integer-threshold crossings
      let cumulative = 0;
      for (let i = 1; i < rowCount; i++) {
        const absVal = absVals[i] || 0;
        const addVal = addVals[i] || 0;

        const cumulativeBefore = (i === 1) ? start : cumulative;
        const cumulativeAfter = cumulativeBefore + addVal;

        const beforeIndex = Math.floor(cumulativeBefore / step);
        const afterIndex = Math.floor(cumulativeAfter / step);
        let hits = Math.max(0, afterIndex - beforeIndex);
        if (hits > 0 && targetIndex + hits > targets.length) hits = Math.max(0, targets.length - targetIndex);
        if (hits > 0) {
          targetIndex += hits;
          if (selections[i] === 'Ano - významnost') {
            selections[i] = 'Ano - významnost; Ano - NPP';
            reasons[i] = reasons[i] ? (reasons[i] + '; NPP') : 'NPP';
          } else {
            selections[i] = hits === 1 ? 'Ano - NPP' : `Ano - NPP (x${hits})`;
            reasons[i] = 'NPP';
          }
          nppCount += hits;
        }

        cumulative = cumulativeAfter;
        const idxVal = cumulative / step;
        out[i] = [absVal, cumulative, idxVal, selections[i], reasons[i]];
      }

      // write out to worksheet to the right of range using batch processing
      const writeRowCount = rowCount;
      const writeColIndex = range.columnIndex + range.columnCount;
      
      // Write header first
      const headerRange = worksheet.getRangeByIndexes(range.rowIndex, writeColIndex, 1, newCols);
      headerRange.values = [out[0]]; // Header row
      headerRange.format.font.bold = true;
      await context.sync();
      
      // Write data in batches to avoid Excel limits
      const dataBatchSize = OUTPUT_BATCH_SIZE; // Use output-optimized batch size
      for (let batchStart = 1; batchStart < out.length; batchStart += dataBatchSize) {
        const batchEnd = Math.min(batchStart + dataBatchSize, out.length);
        const batchData = out.slice(batchStart, batchEnd);
        const batchRange = worksheet.getRangeByIndexes(
          range.rowIndex + batchStart, 
          writeColIndex, 
          batchData.length, 
          newCols
        );
        batchRange.values = batchData;
        await context.sync();
        
        // Show progress for large datasets
        if (out.length > 10000) {
          const progress = Math.round((batchEnd / out.length) * 100);
          showNotification(`Zapisuje se výstup: ${progress}% (${batchEnd}/${out.length} řádků)`);
        }
      }

      await context.sync();

  // show minimal notification; results are written to the worksheet (no panel output)
  const usedSeed = task6Seed || (document.getElementById('task6-seed') && document.getElementById('task6-seed').value) || '';
  showNotification('Generování dokončeno. Výsledky jsou vloženy do listu.' + (usedSeed ? ` Použit seed: ${usedSeed}` : ''));
    });
    return true;
  } catch (e) {
    console.error('Chyba v generateMonetaryRandomWalkTask6:', e);
    showNotification('Chyba při generování mon. random walk: ' + (e && e.message ? e.message : ''));
    return false;
  }
}

// Compute adjusted sum for Task6 according to the exclude-over-significance setting
async function computeTask6AdjustedSum() {
  try {
    const dataRange = document.getElementById('data-range').value;
    const valueColumn = document.getElementById('value-column').value;
    const paramOut = document.getElementById('task6-param-sum-display');
    const countArea = document.getElementById('task6-count-area');
    const countOut = document.getElementById('task6-significant-count-display');

    if (!dataRange || !valueColumn) {
      showNotification('Prosím nejprve vyberte oblast dat a sloupec (úloha 4).');
      if (paramOut) paramOut.value = '';
      if (countArea) countArea.style.display = 'none';
      return;
    }

    // determine whether to exclude over-significance
    const sel = document.getElementById('task6-exclude-over-significance');
    const exclude = sel && sel.value === 'yes';

    // get significance value from storage
    let significance = null;
    try {
      const raw = localStorage.getItem('significanceConfig');
      if (raw) {
        const cfg = JSON.parse(raw);
        if (cfg && cfg.rawValue !== undefined && cfg.rawValue !== null) significance = Number(cfg.rawValue);
      }
    } catch (e) { /* ignore */ }

    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      let addr = dataRange;
      if (addr.indexOf('!') !== -1) addr = addr.split('!').pop();
      addr = addr.replace(/\$/g, '').trim();
      const range = worksheet.getRange(addr);
      range.load(['rowIndex', 'columnIndex', 'rowCount', 'columnCount']);
      await context.sync();

      const absColIndex = getColumnIndex(String(valueColumn).toUpperCase());
      const rangeStartCol = typeof range.columnIndex === 'number' ? range.columnIndex : 0;
      const relColIndex = absColIndex - rangeStartCol;

      if (relColIndex < 0 || relColIndex >= range.columnCount) {
        showNotification('Zadaný sloupec není v rámci vybraného rozsahu dat.');
        if (paramOut) paramOut.value = '';
        if (countArea) countArea.style.display = 'none';
        return;
      }

      let total = 0;
      let significantCount = 0;
      let batchNum = 0;
      for await (const { values } of readRangeInBatches(context, worksheet, range)) {
        batchNum++;
        showNotification(`Zpracovává se dávka ${batchNum}...`);
        for (let i = 1; i < values.length; i++) {
          const raw = values[i][relColIndex];
          const parsed = parseNumberLoose(raw);
          if (parsed === null) continue;
          const absVal = Math.abs(parsed);
          if (exclude && significance !== null && !Number.isNaN(significance)) {
            if (absVal <= significance) {
              total += absVal;
            } else {
              significantCount += 1;
            }
          } else {
            total += absVal;
          }
        }
      }

      // show results according to selection
      if (!exclude) {
        // show Task5 value reference
        // prefer persisted totalSumABS if available
        const displayVal = (totalSumABS !== null && totalSumABS !== undefined) ? totalSumABS : Math.ceil(total);
        if (paramOut) paramOut.value = formatIntegerWithSpaces(Math.round(displayVal));
        if (countArea) countArea.style.display = 'none';
      } else {
        if (paramOut) paramOut.value = formatIntegerWithSpaces(Math.ceil(total));
        if (countArea) countArea.style.display = 'flex';
        if (countOut) countOut.value = String(significantCount);
      }

      showNotification('Upravená suma spočítána.');
    });

  } catch (e) {
    console.error('Chyba při výpočtu upravené sumy:', e);
    showNotification('Chyba při výpočtu upravené sumy.');
  }
}

function handleModalReset() {
  try {
    // Clear storage and reset config
    localStorage.removeItem('reliabilityConfig');
    delete reliabilityConfig.controlRisk;
    delete reliabilityConfig.inherentRisk;
    delete reliabilityConfig.analyticalTests;
    delete reliabilityConfig.controlTests;
    reliabilitySet = false;
  // Also clear modal select elements so they display '<Vyberte>' next time
  const controlEl = document.getElementById("reliability-control-risk");
  const inherentEl = document.getElementById("reliability-inherent-risk");
  const analyticalEl = document.getElementById("reliability-analytical-tests");
  const controlTestsEl = document.getElementById("reliability-control-tests");
  if (controlEl) controlEl.value = '';
  if (inherentEl) inherentEl.value = '';
  if (analyticalEl) analyticalEl.value = '';
  if (controlTestsEl) controlTestsEl.value = '';
    updateReliabilityDisplay();
    const modal = document.getElementById('reliability-modal');
    if (modal) modal.style.display = 'none';
    showNotification('Nastavení bylo resetováno.');
  
  } catch (e) {
    console.error('Chyba při resetování:', e);
    showNotification('Nepodařilo se resetovat nastavení.');
  }
}

// Handler for modal save button
function handleModalSave() {
  try {
    const modal = document.getElementById("reliability-modal");
    const controlRiskEl = document.getElementById("reliability-control-risk");
    const inherentRiskEl = document.getElementById("reliability-inherent-risk");
    const analyticalTestsEl = document.getElementById("reliability-analytical-tests");
    const controlTestsEl = document.getElementById("reliability-control-tests");

    // Allow partial updates: if user leaves a select empty, keep prior value
    const controlRisk = controlRiskEl && controlRiskEl.value ? controlRiskEl.value : reliabilityConfig.controlRisk;
    const inherentRisk = inherentRiskEl && inherentRiskEl.value ? inherentRiskEl.value : reliabilityConfig.inherentRisk;
  const analyticalTests = analyticalTestsEl && analyticalTestsEl.value ? analyticalTestsEl.value : reliabilityConfig.analyticalTests;
  const controlTests = controlTestsEl && controlTestsEl.value ? controlTestsEl.value : reliabilityConfig.controlTests;
  const rmmLevel = document.getElementById('reliability-rmm-level') && document.getElementById('reliability-rmm-level').value ? document.getElementById('reliability-rmm-level').value : reliabilityConfig.rmmLevel;

    // After fallback, require that all fields have some value
  if (!controlRisk || !inherentRisk || !rmmLevel || !analyticalTests || !controlTests) {
      showNotification("Po uložení musí být vyplněna všechna pole (nebo zvolena dříve). Vyplňte chybějící položky.");
      return;
    }

    // Store the selections into reliabilityConfig
    reliabilityConfig.controlRisk = controlRisk;
    reliabilityConfig.inherentRisk = inherentRisk;
    reliabilityConfig.analyticalTests = analyticalTests;
    reliabilityConfig.controlTests = controlTests;
  reliabilityConfig.rmmLevel = rmmLevel;
    reliabilitySet = true;
    // persist
    saveReliabilityToStorage();
    updateReliabilityDisplay();
    // update display with factor value only
    const numeric = resolveFactorFromTable(reliabilityConfig);
    const el = document.getElementById('reliability-factor');
    if (el) {
      if (numeric !== null) {
        el.placeholder = '';
        el.value = String(numeric);
        applyFactorStyle(el, String(numeric));
      } else {
        el.value = '';
        el.placeholder = '<Nevyplněno>';
        applyFactorStyle(el, null);
      }
    }
  
    if (modal) modal.style.display = "none";
  } catch (err) {
    console.error("Chyba při ukládání nastavení spolehlivosti:", err);
    showNotification("Došlo k neočekávané chybě při ukládání.");
  }
}

// Persistence helpers
function saveReliabilityToStorage() {
  try {
    const payload = {
      controlRisk: reliabilityConfig.controlRisk,
      inherentRisk: reliabilityConfig.inherentRisk,
  analyticalTests: reliabilityConfig.analyticalTests,
  controlTests: reliabilityConfig.controlTests,
  rmmLevel: reliabilityConfig.rmmLevel
    };
    localStorage.setItem('reliabilityConfig', JSON.stringify(payload));
  } catch (e) {
    console.warn('Nelze uložit nastavení do localStorage:', e.message);
  }
}

function loadReliabilityFromStorage() {
  try {
    const raw = localStorage.getItem('reliabilityConfig');
    if (!raw) return;
    const obj = JSON.parse(raw);
    if (!obj) return;
    reliabilityConfig.controlRisk = obj.controlRisk || reliabilityConfig.controlRisk;
  reliabilityConfig.inherentRisk = obj.inherentRisk || reliabilityConfig.inherentRisk;
  reliabilityConfig.rmmLevel = obj.rmmLevel || reliabilityConfig.rmmLevel;
  reliabilityConfig.analyticalTests = obj.analyticalTests || reliabilityConfig.analyticalTests;
  reliabilityConfig.controlTests = obj.controlTests || reliabilityConfig.controlTests;
  reliabilitySet = true;
    // after loading, resolve numeric factor from in-code table and show only factor
    const numeric = resolveFactorFromTable(reliabilityConfig);
    const el = document.getElementById('reliability-factor');
    if (el) {
      if (numeric !== null) {
        el.placeholder = '';
        el.value = String(numeric);
        applyFactorStyle(el, String(numeric));
      } else {
        // show descriptive display if no factor available
        updateReliabilityDisplay();
      }
    }
  
  } catch (e) {
    console.warn('Nelze načíst nastavení z localStorage:', e.message);
  }
}

// Resolve numeric factor from the in-code FS_Parametry table
function resolveFactorFromTable(cfg) {
  if (!cfg || !cfg.controlRisk) return null;
  for (const r of fsParams) {
    if (
      r.controlRisk === cfg.controlRisk &&
      r.inherentRisk === cfg.inherentRisk &&
      r.rmmLevel === cfg.rmmLevel &&
      r.analyticalTests === cfg.analyticalTests &&
      r.controlTests === cfg.controlTests
    ) {
      return r.factor;
    }
  }
  return null;
}

// Aktualizace zobrazení faktoru spolehlivosti
function updateReliabilityDisplay() {
  const el = document.getElementById("reliability-factor");
  if (!reliabilitySet) {
    // Use placeholder so it appears with the same muted style as task 3's placeholder
    el.value = "";
    el.placeholder = "<Nevyplněno>";
  // reset styling when not showing factor
  el.style.color = '';
  el.style.fontWeight = '';
    return;
  }

  // Build display from the new selection fields if present
  const control = reliabilityConfig.controlRisk || '-';
  const inherent = reliabilityConfig.inherentRisk || '-';
  const rmm = reliabilityConfig.rmmLevel || '-';
  const analytical = reliabilityConfig.analyticalTests || '-';
  const controls = reliabilityConfig.controlTests || '-';
  const display = `Kontrolní riziko: ${control}, Přirozené riziko: ${inherent}, Hladina RMM: ${rmm}, Analytické testy: ${analytical}, Testy kontrol: ${controls}`;
  el.placeholder = "";
  el.value = display;
  // reset styling when showing descriptive display
  el.style.color = '';
  el.style.fontWeight = '';
}

// Apply styling when factor is a specific error string
function applyFactorStyle(el, factor) {
  if (!el) return;
  if (factor === 'Test kontrol Error') {
    el.style.color = 'red';
    el.style.fontWeight = '700';
  } else {
    el.style.color = '';
    el.style.fontWeight = '';
  }
}

// Funkce pro výběr datového rozsahu
async function selectDataRange() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("address");
      await context.sync();
      
      document.getElementById("data-range").value = range.address;
      console.log(`Vybrán rozsah: ${range.address}`);
  
    });
  } catch (error) {
    console.error("Chyba při výběru rozsahu:", error);
  showNotification("Chyba při výběru rozsahu. Ujistěte se, že máte vybraný rozsah v Excelu.");
  }
}

// Vybere sloupec z aktivní buňky a zapíše jeho označení do #value-column
async function selectValueColumn() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["columnIndex"]);
      await context.sync();

      const colIdx = range.columnIndex; // 0-based
      const colLetter = columnIndexToLetter(colIdx);
      document.getElementById("value-column").value = colLetter;
  console.log(`Vybrán sloupec: ${colLetter}`);
    });
  } catch (error) {
    console.error("Chyba při výběru sloupce:", error);
  showNotification("Chyba při výběru sloupce. Ujistěte se, že máte vybranou buňku v Excelu.");
  }
}

  

// Převod 0-based indexu sloupce na písmenové označení (0 -> A)
function columnIndexToLetter(index) {
  let letter = '';
  let temp = index + 1;
  while (temp > 0) {
    let rem = (temp - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    temp = Math.floor((temp - 1) / 26);
  }
  return letter;
}

// Hlavní funkce pro generování vzorku - optimalizovaná verze
async function generateSample() {
  try {
    const method = document.getElementById("sampling-method").value;
    const dataRange = document.getElementById("data-range").value;
    const valueColumn = document.getElementById("value-column").value;
    
    if (!dataRange || !valueColumn) {
      showNotification("Prosím vyplňte všechna povinná pole.");
      return;
    }
    
    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      // tolerate addresses like 'Sheet1!$A$1:$F$100' by removing sheet and $
      let addr = dataRange;
      if (addr.indexOf('!') !== -1) addr = addr.split('!').pop();
      addr = addr.replace(/\$/g, '').trim();
      const range = worksheet.getRange(addr);
      
      range.load(['rowIndex', 'columnIndex', 'rowCount', 'columnCount']);
      await context.sync();

      const absColIndex = getColumnIndex(String(valueColumn).toUpperCase());
      const rangeStartCol = typeof range.columnIndex === 'number' ? range.columnIndex : 0;
      const relColIndex = absColIndex - rangeStartCol;

      if (relColIndex < 0 || relColIndex >= range.columnCount) {
        showNotification("Zadaný sloupec není v rámci vybraného rozsahu dat.");
        return;
      }

      const totalRows = range.rowCount - 1; // Exclude header row
      
      // Memory safety checks
      if (totalRows > MEMORY_CRITICAL_THRESHOLD) {
        const confirmed = confirm(
          `Varování: Dataset obsahuje ${totalRows.toLocaleString()} řádků. ` +
          `To je velmi velký dataset, který může způsobit problémy s pamětí. ` +
          `Doporučujeme rozdělit data na menší části. Pokračovat?`
        );
        if (!confirmed) {
          showNotification("Operace zrušena uživatelem.");
          return;
        }
      } else if (totalRows > LARGE_DATASET_THRESHOLD) {
        showNotification(
          `Info: Dataset obsahuje ${totalRows.toLocaleString()} řádků. ` +
          `Zpracování může trvat déle. Prosím čekejte...`
        );
      }
      
      showNotification(`Začíná zpracování ${totalRows.toLocaleString()} řádků...`);
      
      // Generování vzorku podle zvolené metody pomocí streaming algoritmu
      let sample;
      if (method === "monetary-random-walk") {
        sample = await generateMonetaryRandomWalkStreaming(context, worksheet, range, relColIndex, totalRows);
      } else {
        sample = await generateRandomNumberSampleStreaming(context, worksheet, range, relColIndex, totalRows);
      }
      
      // Zobrazení výsledků
      displayResults(sample, totalRows + 1); // +1 for header
      
      // Zvýraznění vybraných řádků
      await highlightSampleRows(context, worksheet, dataRange, sample);
      
      showNotification(`Vzorkování dokončeno! Vybráno ${sample.length} vzorků z ${totalRows.toLocaleString()} řádků.`);
    });
  } catch (error) {
    console.error("Chyba při generování vzorku:", error);
    
    // Better error messages for common issues
    let errorMessage = "Chyba při generování vzorku: ";
    if (error.message && error.message.includes("memory")) {
      errorMessage += "Nedostatek paměti. Zkuste rozdělit data na menší části.";
    } else if (error.message && error.message.includes("timeout")) {
      errorMessage += "Operace trvala příliš dlouho. Zkuste menší dataset.";
    } else {
      errorMessage += (error && error.message ? error.message : 'Neznámá chyba');
    }
    
    showNotification(errorMessage, 8000); // Longer display for errors
  }
}

// Převod označení sloupce na index
function getColumnIndex(columnRef) {
  if (!isNaN(columnRef)) {
    return parseInt(columnRef) - 1;
  }
  
  let result = 0;
  for (let i = 0; i < columnRef.length; i++) {
    result = result * 26 + (columnRef.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  return result - 1;
}
 // Helper to write rows in batches with improved memory management for output
 async function writeRangeInBatches(context, worksheet, startRowIdx, startColIdx, rows, batchSize = OUTPUT_BATCH_SIZE) {
   for (let offset = 0; offset < rows.length; offset += batchSize) {
     const batchRows = rows.slice(offset, offset + batchSize);
     const batchRowCount = batchRows.length;
     
     try {
       const writeRange = worksheet.getRangeByIndexes(startRowIdx + offset, startColIdx, batchRowCount, batchRows[0].length);
       writeRange.values = batchRows;
       await context.sync();
       
       // Progress for large output operations
       if (rows.length > 5000 && offset % (batchSize * 5) === 0) {
         const progress = Math.round(((offset + batchRowCount) / rows.length) * 100);
         showNotification(`Zapisuje se ${progress}% výsledků...`);
       }
     } catch (error) {
       console.error(`Chyba při zápisu dávky na pozici ${offset}:`, error);
       throw new Error(`Chyba při zápisu výsledků na řádku ${startRowIdx + offset}: ${error.message}`);
     }
   }
 }

 // Helper to read rows in batches with improved memory management
 async function* readRangeInBatches(context, worksheet, range, batchSize = BATCH_SIZE) {
   const totalRows = range.rowCount;
   const startRow = range.rowIndex;
   const startCol = range.columnIndex;
   const colCount = range.columnCount;
   
   // Memory warning for very large datasets
   if (totalRows > 50000) {
     showNotification(`Upozornění: Zpracovává se ${totalRows.toLocaleString()} řádků. Prosím čekejte...`);
   }
   
   for (let offset = 0; offset < totalRows; offset += batchSize) {
     const rows = Math.min(batchSize, totalRows - offset);
     
     try {
       const subRange = worksheet.getRangeByIndexes(startRow + offset, startCol, rows, colCount);
       subRange.load('values');
       await context.sync();
       
       yield { values: subRange.values, startRow: startRow + offset };
       
       // Force garbage collection hint for large datasets
       if (totalRows > 100000 && offset % (batchSize * 10) === 0) {
         showNotification(`Zpracováno ${Math.round((offset / totalRows) * 100)}% - ${(offset + rows).toLocaleString()}/${totalRows.toLocaleString()} řádků...`);
         // Give browser a chance to garbage collect
         await new Promise(resolve => setTimeout(resolve, 10));
       }
     } catch (error) {
       console.error(`Chyba při čtení dávky na pozici ${offset}:`, error);
       throw new Error(`Chyba při čtení dat na řádku ${startRow + offset}: ${error.message}`);
     }
   }
 }


// Optimalizovaná implementace náhodné peněžní procházky - streaming verze
async function generateMonetaryRandomWalkStreaming(context, worksheet, range, columnIndex, totalRows) {
  // Nejdříve potřebujeme spočítat celkovou sumu pro výpočet intervalu
  showNotification("Počítá se celková suma...");
  let totalValue = 0;
  let batchNum = 0;
  
  for await (const { values } of readRangeInBatches(context, worksheet, range, BATCH_SIZE)) {
    batchNum++;
    showNotification(`Počítá se suma - dávka ${batchNum}...`);
    const startIdx = batchNum === 1 ? 1 : 0; // Skip header in first batch
    
    for (let i = startIdx; i < values.length; i++) {
      const raw = values[i][columnIndex];
      const parsed = parseNumberLoose(raw);
      const value = parsed !== null ? parsed : 0;
      totalValue += Math.abs(value);
    }
  }

  const sampleSize = calculateSampleSize(totalValue);
  if (!sampleSize || sampleSize < 1) {
    showNotification("Vypočtená velikost vzorku je neplatná.");
    return [];
  }

  const interval = totalValue / sampleSize;
  const sample = [];
  let currentSum = 0;
  const rng = getTask6Rng();
  let nextTarget = (typeof rng === 'function' ? rng() : Math.random()) * interval;
  let globalRowIndex = 0; // Global row index (including header)
  
  showNotification(`Generuje se vzorek ${sampleSize} položek z celkové sumy ${totalValue.toLocaleString()}...`);
  batchNum = 0;
  
  // Druhý průchod pro výběr vzorku
  for await (const { values } of readRangeInBatches(context, worksheet, range, BATCH_SIZE)) {
    batchNum++;
    const startIdx = batchNum === 1 ? 1 : 0; // Skip header in first batch
    
    showNotification(`Zpracování vzorku - dávka ${batchNum}/${Math.ceil(totalRows / BATCH_SIZE)}...`);
    
    for (let i = startIdx; i < values.length; i++) {
      globalRowIndex++;
      const raw = values[i][columnIndex];
      const parsed = parseNumberLoose(raw);
      const value = parsed !== null ? Math.abs(parsed) : 0;
      currentSum += value;
      
      if (currentSum >= nextTarget && sample.length < sampleSize) {
        sample.push({
          rowIndex: globalRowIndex,
          rowData: values[i],
          value: raw,
          cumulativeValue: currentSum
        });
        nextTarget += interval;
      }
      
      // Early exit if we have enough samples
      if (sample.length >= sampleSize) {
        showNotification(`Vzorkování dokončeno - nalezeno ${sample.length} vzorků.`);
        return sample;
      }
    }
  }
  
  return sample;
}

// Optimalizovaná implementace náhodného generátoru čísel - streaming verze
async function generateRandomNumberSampleStreaming(context, worksheet, range, columnIndex, totalRows) {
  // Nejdříve potřebujeme spočítat celkovou sumu pro výpočet velikosti vzorku
  showNotification("Počítá se celková suma...");
  let totalValue = 0;
  let batchNum = 0;
  
  for await (const { values } of readRangeInBatches(context, worksheet, range, BATCH_SIZE)) {
    batchNum++;
    showNotification(`Počítá se suma - dávka ${batchNum}...`);
    const startIdx = batchNum === 1 ? 1 : 0; // Skip header in first batch
    
    for (let i = startIdx; i < values.length; i++) {
      const raw = values[i][columnIndex];
      const parsed = parseNumberLoose(raw);
      const value = parsed !== null ? parsed : 0;
      totalValue += Math.abs(value);
    }
  }

  const sampleSize = calculateSampleSize(totalValue);
  if (!sampleSize || sampleSize < 1) {
    showNotification("Vypočtená velikost vzorku je neplatná.");
    return [];
  }

  // Předem vybereme náhodné indexy řádků (0-based, bez hlavičky)
  const selectedIndices = new Set();
  const rng = getTask6Rng();
  
  while (selectedIndices.size < Math.min(sampleSize, totalRows)) {
    const randomIndex = Math.floor((typeof rng === 'function' ? rng() : Math.random()) * totalRows);
    selectedIndices.add(randomIndex);
  }
  
  showNotification(`Generuje se vzorek ${selectedIndices.size} položek z ${totalRows.toLocaleString()} řádků...`);
  
  const sample = [];
  let globalRowIndex = 0; // Global row index (0-based, excluding header)
  batchNum = 0;
  
  // Průchod daty a výběr vzorku podle předem vybraných indexů
  for await (const { values } of readRangeInBatches(context, worksheet, range, BATCH_SIZE)) {
    batchNum++;
    const startIdx = batchNum === 1 ? 1 : 0; // Skip header in first batch
    
    showNotification(`Zpracování vzorku - dávka ${batchNum}/${Math.ceil(totalRows / BATCH_SIZE)}...`);
    
    for (let i = startIdx; i < values.length; i++) {
      if (selectedIndices.has(globalRowIndex)) {
        sample.push({
          rowIndex: globalRowIndex + 1, // +1 because we want 1-based indexing relative to header
          rowData: values[i],
          value: values[i][columnIndex]
        });
      }
      globalRowIndex++;
      
      // Early exit if we have all samples
      if (sample.length >= selectedIndices.size) {
        showNotification(`Vzorkování dokončeno - nalezeno ${sample.length} vzorků.`);
        return sample;
      }
    }
  }
  
  return sample;
}

// Výpočet velikosti vzorku na základě faktoru spolehlivosti
// Calculate sample size using TotalSum * factor / materiality when totalSum is provided.
// If totalSum is not provided or materiality/factor missing, falls back to the previous statistical method.
function calculateSampleSize(totalSum) {
  // Try to read factor from the reliability display first
  let factorNum = null;
  try {
    const factorEl = document.getElementById('reliability-factor');
    if (factorEl && factorEl.value) {
      const parsed = parseNumberLoose(factorEl.value);
      if (parsed !== null) factorNum = parsed;
    }
  } catch (e) {}

  // If not numeric, try to resolve from table mapping
  if (factorNum === null || Number.isNaN(factorNum)) {
    const resolved = resolveFactorFromTable(reliabilityConfig);
    if (typeof resolved === 'number' && !Number.isNaN(resolved)) factorNum = resolved;
  }

  // Try to read materiality from stored significance config
  let materiality = null;
  try {
    const raw = localStorage.getItem('significanceConfig');
    if (raw) {
      const cfg = JSON.parse(raw);
      if (cfg && typeof cfg.rawValue === 'number') materiality = cfg.rawValue;
      // if rawValue stored as string, try to parse
      else if (cfg && cfg.rawValue) {
        const m = parseNumberLoose(cfg.rawValue);
        if (m !== null) materiality = m;
      }
    }
  } catch (e) {}

  // If totalSum provided and factor & materiality available, use the absolute-sum formula
  if (typeof totalSum === 'number' && !Number.isNaN(totalSum) && factorNum && materiality && materiality > 0) {
    const raw = (totalSum * factorNum) / materiality;
    const n = Math.ceil(raw);
    return Math.max(1, n);
  }

  // Fallback: previous simple statistical method using reliabilityConfig
  try {
    const { confidence, expectedError, populationSize } = reliabilityConfig;
    const z = confidence === 95 ? 1.96 : confidence === 99 ? 2.58 : 1.645;
    const p = 0.5;
    const e = expectedError / 100;
    let n = (z * z * p * (1 - p)) / (e * e);
    if (populationSize > 0) {
      n = n / (1 + (n - 1) / populationSize);
    }
    return Math.ceil(n);
  } catch (e) {
    return 1;
  }
}

// Helper to compute absolute sum for a column from a 2D values array (skip header row)
function getTotalAbsoluteSum(values, columnIndex) {
  // Expect values to be a 2D array as returned by Excel range.values
  // Sum absolute values starting from second row (skip header)
  if (!Array.isArray(values) || values.length <= 1) return 0;
  return values.slice(1).reduce((sum, row) => {
    const raw = row && row.length > columnIndex ? row[columnIndex] : null;
    const parsed = parseNumberLoose(raw);
    if (parsed === null) return sum;
    return sum + Math.abs(parsed);
  }, 0);
}

// Robust number parser: accepts '1 234,56', '1234.56', '1,234.56' and returns Number or null
function parseNumberLoose(input) {
  if (input === null || input === undefined) return null;
  let s = String(input).trim();
  if (s === '') return null;
  // remove non-breaking spaces and regular spaces
  s = s.replace(/\s+/g, '');
  // If contains comma but no dot, assume comma is decimal separator
  if (s.indexOf(',') !== -1 && s.indexOf('.') === -1) {
    s = s.replace(',', '.');
  } else {
    // remove commas used as thousand separators
    s = s.replace(/,/g, '');
  }
  // remove any other non-digit/.- characters
  s = s.replace(/[^0-9.\-]/g, '');
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

// Persist last computed TotalSumABS
function saveTotalSumAbsToStorage(val) {
  try {
    localStorage.setItem('TotalSumABS', JSON.stringify(val));
  } catch (e) {
    console.warn('Nelze uložit TotalSumABS:', e && e.message ? e.message : e);
  }
}

function loadTotalSumAbsFromStorage() {
  try {
    const raw = localStorage.getItem('TotalSumABS');
    if (!raw) return;
    const v = JSON.parse(raw);
    if (v !== null && v !== undefined) {
      totalSumABS = v;
      const outEl = document.getElementById('task5-result-display');
      if (outEl) outEl.value = formatIntegerWithSpaces(Math.round(totalSumABS));
    }
  } catch (e) {
    console.warn('Nelze načíst TotalSumABS:', e && e.message ? e.message : e);
  }
}

// Persist/load Task6 exclude-over-significance choice
function saveTask6ExcludeSetting(val) {
  try {
    localStorage.setItem('task6ExcludeOverSignificance', String(val));
  } catch (e) {
    console.warn('Nelze uložiť volbu Task6 do storage:', e && e.message ? e.message : e);
  }
}

function saveTask6PrintTarget(val) {
  try { localStorage.setItem('task6PrintTarget', String(val)); } catch (e) { console.warn('Nelze uložit tisk target:', e); }
}

function loadTask6PrintTarget() {
  try {
    const raw = localStorage.getItem('task6PrintTarget');
    const sel = document.getElementById('task6-print-target');
    if (!sel) return;
    sel.value = raw || '';
  } catch (e) { console.warn('Nelze načíst tisk target:', e); }
}

// Seed persistence for Task6
function saveTask6Seed(val) {
  try { localStorage.setItem('task6Seed', String(val)); task6Seed = val; } catch (e) { console.warn('Nelze uložit seed:', e); }
}

function loadTask6Seed() {
  try {
    const raw = localStorage.getItem('task6Seed');
    const el = document.getElementById('task6-seed');
    if (raw && el) { el.value = raw; task6Seed = raw; }
  } catch (e) { console.warn('Nelze načíst seed:', e); }
}

// Small deterministic PRNG helpers: xfnv1a hash -> mulberry32
function xfnv1aHash(str) {
  let h = 2166136261 >>> 0;
  for (let i = 0; i < str.length; i++) {
    h ^= str.charCodeAt(i);
    h = Math.imul(h, 16777619) >>> 0;
  }
  h += h << 13; h ^= h >>> 7; h += h << 3; h ^= h >>> 17; h += h << 5;
  return h >>> 0;
}

function mulberry32(a) {
  return function() {
    a |= 0;
    a = (a + 0x6D2B79F5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

// Return RNG function or null to indicate fallback to Math.random
function getTask6Rng() {
  try {
    const seedVal = task6Seed || (document.getElementById('task6-seed') && document.getElementById('task6-seed').value);
    if (!seedVal && seedVal !== 0) return null;
    const s = String(seedVal);
    const h = xfnv1aHash(s);
    return mulberry32(h);
  } catch (e) { return null; }
}

// Create a short random seed string (timestamp + random suffix)
function makeRandomSeed() {
  const t = Date.now().toString(36);
  const r = Math.floor(Math.random() * 1e6).toString(36);
  return `${t}-${r}`;
}

function loadTask6ExcludeSetting() {
  try {
    const raw = localStorage.getItem('task6ExcludeOverSignificance');
    const sel = document.getElementById('task6-exclude-over-significance');
    if (!sel) return;
    if (!raw) { sel.value = ''; return; }
    sel.value = raw;
  } catch (e) {
    console.warn('Nelze načíst volbu Task6 z storage:', e && e.message ? e.message : e);
  }
}

  

  

// Zobrazení výsledků
function displayResults(sample, totalRows) {
  const resultsDiv = document.getElementById("results");
  const sampleInfo = document.getElementById("sample-info");
  const sampleList = document.getElementById("sample-list");
  
  sampleInfo.innerHTML = `
    <strong>Celkem řádků:</strong> ${totalRows}<br>
    <strong>Velikost vzorku:</strong> ${sample.length}<br>
    <strong>Procento vzorku:</strong> ${((sample.length / totalRows) * 100).toFixed(2)}%
  `;
  
  let listHtml = "<strong>Vybrané řádky:</strong><br>";
  sample.forEach((item, index) => {
    listHtml += `${index + 1}. Řádek ${item.rowIndex + 1}, Hodnota: ${item.value}<br>`;
  });
  
  sampleList.innerHTML = listHtml;
  resultsDiv.style.display = "block";
}

// Zvýraznění vybraných řádků v Excelu
async function highlightSampleRows(context, worksheet, dataRange, sample) {
  try {
    // Nejdříve vyčistíme předchozí zvýraznění
    const fullRange = worksheet.getRange(dataRange);
    fullRange.format.fill.clear();
    
    // Zvýrazníme vybrané řádky
    for (const item of sample) {
      const rowAddress = dataRange.split(':')[0].replace(/\d+/, (item.rowIndex + 1).toString()) + 
                        ':' + 
                        dataRange.split(':')[1].replace(/\d+/, (item.rowIndex + 1).toString());
      
      const rowRange = worksheet.getRange(rowAddress);
      rowRange.format.fill.color = "yellow";
    }
    
    await context.sync();
  } catch (error) {
    console.log("Varování: Nepodařilo se zvýraznit řádky:", error.message);
  }
}

// Zachování původní run funkce pro kompatibilitu
export async function run() {
  await generateSample();
}

// Simple notification helper (replaces unsupported alert())
function showNotification(message, timeout = 4000) {
  try {
    const el = document.getElementById('notification');
    if (!el) return;
    el.textContent = message;
    el.style.display = 'block';
    setTimeout(() => { el.style.display = 'none'; }, timeout);
  } catch (e) {
    // last resort: log to console
    console.log('Notifikace:', message);
  }
}

// -------------------- Significance (Hladina Významnosti) --------------------
// significanceConfig structure: { type: string, rawValue: number|null, justification: string }
function openSignificanceDialog() {
  const modal = document.getElementById('significance-modal');
  const typeEl = document.getElementById('significance-type');
  const valueEl = document.getElementById('significance-value');
  const justEl = document.getElementById('significance-justification');
  if (!modal || !typeEl || !valueEl || !justEl) {
    showNotification('Dialog hladiny významnosti není dostupný.');
    return;
  }
  // load stored config
  const raw = localStorage.getItem('significanceConfig');
  let cfg = null;
  try { cfg = raw ? JSON.parse(raw) : null; } catch (e) { cfg = null; }
  if (cfg) {
    typeEl.value = cfg.type || '';
    // rawValue saved as number -> show unformatted in edit field
    valueEl.value = (cfg.rawValue !== undefined && cfg.rawValue !== null) ? String(cfg.rawValue) : '';
    justEl.value = cfg.justification || '';
  } else {
    typeEl.value = '';
    valueEl.value = '';
    justEl.value = '';
  }
  // Apply rule: if Prováděcí then justification readonly and set to 'NEVYPLŇOVAT'
  if (typeEl.value === 'Prováděcí') {
    justEl.value = 'NEVYPLŇOVAT';
    justEl.readOnly = true;
  } else {
    justEl.readOnly = false;
    if (justEl.value === 'NEVYPLŇOVAT') justEl.value = '';
  }

  // Wire dynamic behavior inside modal: when type changes adjust justification field
  typeEl.onchange = () => {
    if (typeEl.value === 'Prováděcí') {
      justEl.value = 'NEVYPLŇOVAT';
      justEl.readOnly = true;
    } else {
      if (justEl.value === 'NEVYPLŇOVAT') justEl.value = '';
      justEl.readOnly = false;
    }
  };

  // Live formatting for the numeric input (group thousands while typing)
  const liveFormatFn = (ev) => {
    formatIntegerInputWithCursor(valueEl, ev);
  };
  valueEl.removeEventListener('input', liveFormatFn);
  valueEl.addEventListener('input', liveFormatFn);

  modal.style.display = 'flex';
}

function handleSignificanceSave() {
  try {
    const typeEl = document.getElementById('significance-type');
    const valueEl = document.getElementById('significance-value');
    const justEl = document.getElementById('significance-justification');
    if (!typeEl || !valueEl || !justEl) return;

    const type = typeEl.value || '';
    let rawInput = valueEl.value ? valueEl.value.trim() : '';
    let justification = justEl.value ? justEl.value.trim() : '';

    // Basic validation: type required
    if (!type) {
      showNotification('Prosím vyberte typ významnosti.');
      return;
    }

    // If type is not Prováděcí, justification is required
    if (type !== 'Prováděcí' && !justification) {
      showNotification('Pokud není zvoleno "Prováděcí", je potřeba zadat zdůvodnění.');
      return;
    }

    // Normalize numeric value: remove percent, non-digit etc., then parse integer and take absolute
    let rawNumber = null;
    if (rawInput) {
      rawInput = rawInput.replace(/%/g, '');
      const digits = rawInput.replace(/[^0-9\-]/g, '');
      if (digits !== '' && digits !== '-' && digits !== '+') {
        rawNumber = parseInt(digits, 10);
        if (Number.isNaN(rawNumber)) rawNumber = null;
        else rawNumber = Math.abs(rawNumber);
      }
    }

    // Require value to be present (as integer)
    if (rawNumber === null) {
      showNotification('Prosím zadejte hodnotu hladiny významnosti jako celé číslo.');
      return;
    }

    if (type === 'Prováděcí') justification = 'NEVYPLŇOVAT';

    const cfg = { type, rawValue: rawNumber, justification };
    localStorage.setItem('significanceConfig', JSON.stringify(cfg));

    // Update main panel displays
    const typeDisplay = document.getElementById('significance-type-display');
    const valueDisplay = document.getElementById('significance-value-display');
    const justDisplay = document.getElementById('significance-justification-display');
    if (typeDisplay) typeDisplay.value = type;
    if (valueDisplay) valueDisplay.value = formatIntegerWithSpaces(rawNumber);
    if (justDisplay) justDisplay.value = justification;

  

    const modal = document.getElementById('significance-modal');
    if (modal) modal.style.display = 'none';
    showNotification('Hladina významnosti uložena.');
  } catch (e) {
    console.error('Chyba při ukládání hladiny významnosti:', e);
    showNotification('Nelze uložit hladinu významnosti.');
  }
}

function handleSignificanceReset() {
  try {
    localStorage.removeItem('significanceConfig');
    const typeDisplay = document.getElementById('significance-type-display');
    const valueDisplay = document.getElementById('significance-value-display');
    const justDisplay = document.getElementById('significance-justification-display');
    if (typeDisplay) { typeDisplay.value = ''; typeDisplay.placeholder = 'Typ významnosti'; }
    if (valueDisplay) { valueDisplay.value = ''; valueDisplay.placeholder = 'Hodnota (bez deset.)'; }
    if (justDisplay) { justDisplay.value = ''; justDisplay.placeholder = 'Zdůvodnění / stav'; }
  
    const modal = document.getElementById('significance-modal');
    if (modal) modal.style.display = 'none';
    showNotification('Hladina významnosti byla resetována.');
  } catch (e) {
    console.error('Chyba při resetování hladiny významnosti:', e);
    showNotification('Nelze resetovat hladinu významnosti.');
  }
}

function loadSignificanceFromStorage() {
  try {
    const raw = localStorage.getItem('significanceConfig');
    let cfg = null;
    try { cfg = raw ? JSON.parse(raw) : null; } catch (e) { cfg = null; }
    const typeDisplay = document.getElementById('significance-type-display');
    const valueDisplay = document.getElementById('significance-value-display');
    const justDisplay = document.getElementById('significance-justification-display');
    if (cfg) {
      if (typeDisplay) typeDisplay.value = cfg.type || '';
      if (valueDisplay) valueDisplay.value = (cfg.rawValue !== undefined && cfg.rawValue !== null) ? formatIntegerWithSpaces(cfg.rawValue) : '';
      if (justDisplay) justDisplay.value = cfg.justification || '';
    } else {
      if (typeDisplay) { typeDisplay.value = ''; typeDisplay.placeholder = 'Typ významnosti'; }
      if (valueDisplay) { valueDisplay.value = ''; valueDisplay.placeholder = 'Hodnota (bez deset.)'; }
      if (justDisplay) { justDisplay.value = ''; justDisplay.placeholder = 'Zdůvodnění / stav'; }
    }
  } catch (e) {
    console.warn('Nelze načíst hladinu významnosti ze storage:', e.message);
  }
}

function formatIntegerWithSpaces(n) {
  if (n === null || n === undefined) return '';
  const s = String(Math.round(n));
  return s.replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
}

// Format input element's value as integer with spaces while trying to preserve caret position
function formatIntegerInputWithCursor(inputEl, ev) {
  if (!inputEl) return;
  const old = inputEl.value;
  const selStart = inputEl.selectionStart;
  // extract digits and sign
  let cleaned = old.replace(/%/g, '').replace(/[^0-9\-]/g, '');
  if (cleaned === '' || cleaned === '-' || cleaned === '+') {
    inputEl.value = cleaned;
    return;
  }
  const asNum = Math.abs(parseInt(cleaned, 10));
  if (Number.isNaN(asNum)) {
    inputEl.value = '';
    return;
  }
  const formatted = String(asNum).replace(/\B(?=(\d{3})+(?!\d))/g, ' ');

  // compute new caret position: estimate number of non-digit chars before original caret
  let digitsBefore = 0;
  for (let i = 0; i < selStart; i++) if (/[0-9]/.test(old[i])) digitsBefore++;
  // map digitsBefore into formatted position
  let newPos = 0;
  let digitsSeen = 0;
  while (newPos < formatted.length && digitsSeen < digitsBefore) {
    if (/[0-9]/.test(formatted[newPos])) digitsSeen++;
    newPos++;
  }

  inputEl.value = formatted;
  // set caret
  try { inputEl.setSelectionRange(newPos, newPos); } catch (e) {}
}

// Task 5: compute sum of absolute values for chosen column within chosen range
async function computeTask5Sum() {
  // declare outEl in outer scope so the catch block can clear it on error
  let outEl = null;
  try {
    const dataRange = document.getElementById('data-range').value;
    const valueColumn = document.getElementById('value-column').value;
    outEl = document.getElementById('task5-result-display');

    if (!dataRange || !valueColumn) {
      showNotification('Prosím nejprve vyberte oblast dat a sloupec (úloha 4).');
      if (outEl) outEl.value = '';
      return;
    }

    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      // sanitize address (remove sheet name and $)
      let addr = dataRange;
      if (addr.indexOf('!') !== -1) addr = addr.split('!').pop();
      addr = addr.replace(/\$/g, '').trim();

      const range = worksheet.getRange(addr);
      range.load(['rowIndex', 'columnIndex', 'rowCount', 'columnCount']);
      await context.sync();

      const absColIndex = getColumnIndex(String(valueColumn).toUpperCase());
      const rangeStartCol = typeof range.columnIndex === 'number' ? range.columnIndex : 0;
      const relColIndex = absColIndex - rangeStartCol;

      if (relColIndex < 0 || relColIndex >= range.columnCount) {
        showNotification('Zadaný sloupec není v rámci vybraného rozsahu dat.');
        if (outEl) outEl.value = '';
        return;
      }

      let total = 0;
      let batchNum = 0;
      for await (const { values } of readRangeInBatches(context, worksheet, range)) {
        batchNum++;
        showNotification(`Zpracovává se dávka ${batchNum}...`);
        total += getTotalAbsoluteSum(values, relColIndex);
      }

      // store and display as integer rounded up to units
      const totalCeil = Math.ceil(total);
      if (outEl) outEl.value = formatIntegerWithSpaces(totalCeil);
      // persist value (integer)
      totalSumABS = totalCeil;
      try { saveTotalSumAbsToStorage(totalSumABS); } catch (e) {}
      showNotification('Suma byla spočítána.');
    });
  } catch (e) {
  console.error('Chyba při výpočtu Task5 sumy:', e && e.stack ? e.stack : e);
  const msg = (e && e.message) ? e.message : String(e);
  if (outEl) outEl.value = '';
  showNotification('Chyba při výpočtu sumy: ' + msg);
  }
}

// Compute sample size using persisted TotalSumABS, reliability factor and significance
function computeTask5Size() {
  try {
    const outEl = document.getElementById('task5-size-display');

    // Read factor from the display (task 2) or resolve from table
    let factorNum = null;
    const factorEl = document.getElementById('reliability-factor');
    if (factorEl && factorEl.value) {
      const p = parseNumberLoose(factorEl.value);
      if (p !== null) factorNum = p;
    }
    if (factorNum === null || Number.isNaN(factorNum)) {
      const resolved = resolveFactorFromTable(reliabilityConfig);
      if (typeof resolved === 'number' && !Number.isNaN(resolved)) factorNum = resolved;
    }

    // Read significance
    let significance = null;
    try {
      const raw = localStorage.getItem('significanceConfig');
      if (raw) {
        const cfg = JSON.parse(raw);
        if (cfg && typeof cfg.rawValue === 'number') significance = cfg.rawValue;
        else if (cfg && cfg.rawValue) {
          const m = parseNumberLoose(cfg.rawValue);
          if (m !== null) significance = m;
        }
      }
    } catch (e) {}

    // Read TotalSumABS from memory/storage or UI
    let total = totalSumABS;
    if ((total === null || total === undefined) && document.getElementById('task5-result-display')) {
      const txt = document.getElementById('task5-result-display').value;
      const parsed = parseNumberLoose(txt);
      if (parsed !== null) total = parsed;
    }

    if (factorNum === null || significance === null || total === null || significance === 0) {
      if (outEl) outEl.value = '';
      showNotification('Pro výpočet velikosti vzorku musí být nastaven faktor, významnost a spočítaná suma.');
      return;
    }

    const raw = (factorNum * total) / significance;
    const n = Math.ceil(raw);
    if (outEl) outEl.value = String(n);
    showNotification('Velikost vzorku byla spočítána.');
  } catch (e) {
    console.error('Chyba při výpočtu velikosti vzorku:', e && e.stack ? e.stack : e);
    const msg = (e && e.message) ? e.message : String(e);
    const outEl = document.getElementById('task5-size-display');
    if (outEl) outEl.value = '';
    showNotification('Chyba při výpočtu velikosti vzorku: ' + msg);
  }
}

// Show/hide override inputs when user chooses whether to use calculated size
function handleTask5UseChoice(ev) {
  try {
    const sel = document.getElementById('task5-use-calculated');
    const area = document.getElementById('task5-override-area');
    const overrideSize = document.getElementById('task5-override-size');
    const overrideReason = document.getElementById('task5-override-reason');
    if (!sel || !area) return;
    if (sel.value === 'no') {
      area.style.display = 'flex';
      area.style.flexDirection = 'column';
    } else {
      // hide and clear
      area.style.display = 'none';
      if (overrideSize) overrideSize.value = '';
      if (overrideReason) overrideReason.value = '';
    }
  } catch (e) {
    console.error('Error handling Task5 use choice:', e);
  }
}

function confirmRequestedSize() {
  try {
    const useSel = document.getElementById('task5-use-calculated');
    const overrideSizeEl = document.getElementById('task5-override-size');
    const overrideReasonEl = document.getElementById('task5-override-reason');
    const infoEl = document.getElementById('task5-confirm-info');

    if (!useSel) {
      showNotification('Vyberte prosím, zda chcete použít vypočtenou velikost.');
      return;
    }

    if (useSel.value === 'yes') {
      // Use the value currently displayed in the size field (task5-size-display)
      const sizeEl = document.getElementById('task5-size-display');
      if (!sizeEl || !sizeEl.value) {
        showNotification('Neexistuje vypočtená hodnota. Nejprve spočítejte velikost.');
        return;
      }
      const parsed = parseNumberLoose(sizeEl.value);
      if (parsed === null) {
        showNotification('Nelze načíst vypočtenou velikost z pole.');
        return;
      }
      const requested = Math.ceil(parsed);
      requestedSampleSize = requested;
      requestedSampleReason = 'Použito vypočtené n';
  if (infoEl) infoEl.textContent = '';
  // populate Task 6 final displays
  const task6Size = document.getElementById('task6-final-size-display');
  if (task6Size) task6Size.value = String(requestedSampleSize);
  const task6Method = document.getElementById('task6-method-display');
  const samplingMethodEl = document.getElementById('sampling-method');
  if (task6Method && samplingMethodEl) task6Method.value = samplingMethodEl.options[samplingMethodEl.selectedIndex].text;
  // (override flag/reason are shown in Task 5 only)
      showNotification('Velikost vzorku potvrzena (použito vypočtené).');
      // reflect in UI (ensure integer display)
      if (sizeEl) sizeEl.value = String(requestedSampleSize);
      return;
    }

    if (useSel.value === 'no') {
      // Use override value
      if (!overrideSizeEl || !overrideSizeEl.value) {
        showNotification('Zadejte prosím požadovanou velikost.');
        return;
      }
      const parsed = parseNumberLoose(overrideSizeEl.value);
      if (parsed === null || parsed <= 0) {
        showNotification('Žádaná velikost musí být kladné číslo.');
        return;
      }
      const requested = Math.ceil(parsed);
      requestedSampleSize = requested;
      requestedSampleReason = overrideReasonEl ? (overrideReasonEl.value || '') : '';
  if (infoEl) infoEl.textContent = '';

  // START OF MODIFIED LOGIC
  // Only update Task 6's "Wanted sample size" display
  const task6Size = document.getElementById('task6-final-size-display');
  if (task6Size) task6Size.value = String(requestedSampleSize);
  
  // DO NOT update the main size display in Task 5
  showNotification('Žádaná velikost byla potvrzena a aplikována pouze do Úkolu 6.');
  return;
    }

    showNotification('Vyberte prosím možnost Ano nebo Ne.');
  } catch (e) {
    console.error('Chyba při potvrzení požadované velikosti:', e);
    showNotification('Chyba při potvrzení velikosti.');
  }
}

// Helper to safely get value from an element
function _get_value(id) {
  const el = document.getElementById(id);
  if (el) {
    if (el.tagName === 'SELECT') {
      try { return el.options[el.selectedIndex].text; } catch (e) { return ''; }
    }
    return el.value || '';
  }
  return '';
}

function _get_selected_value(id) {
  const el = document.getElementById(id);
  return el ? el.value : '';
}

// Formats parameters into a 2D array for writing to the worksheet
function _format_params(params) {
  const out = [];
  out.push(['Parametry vzorkovače', '']);
  for (const key in params) {
    if (Object.hasOwnProperty.call(params, key)) {
      const value = params[key];
      out.push([key, value === null || value === undefined ? '' : value]);
    }
  }
  return out;
}

// Prints parameters to the specified location
async function printParameters(params, target) {
  try {
    await Excel.run(async (context) => {
      const formatted = _format_params(params);
      let worksheet, startRow = 0, startCol = 0;
      
      if (target === 'new-sheet') {
        const newSheet = context.workbook.worksheets.add("ParametryVzorku");
        newSheet.activate();
        worksheet = newSheet;
        startRow = 0;
        startCol = 0;
      } else if (target === 'above') {
        const dataRange = document.getElementById('data-range').value;
        if (!dataRange) {
          showNotification('Oblast dat není vybrána. Parametry budou vloženy na nový list.');
          const newSheet = context.workbook.worksheets.add("ParametryVzorku");
          newSheet.activate();
          worksheet = newSheet;
          startRow = 0;
          startCol = 0;
        } else {
          let addr = dataRange;
          if (addr.indexOf('!') !== -1) addr = addr.split('!').pop();
          addr = addr.replace(/\$/g, '').trim();
          const range = context.workbook.worksheets.getActiveWorksheet().getRange(addr);
          range.load(['rowIndex']);
          await context.sync();
          startRow = range.rowIndex;
          startCol = 0;
          worksheet = context.workbook.worksheets.getActiveWorksheet();
          const usedRange = worksheet.getUsedRange();
          usedRange.load('columnCount');
          await context.sync();
          const columnCount = usedRange.columnCount;
          worksheet.getRangeByIndexes(startRow, 0, formatted.length, columnCount).insert(Excel.InsertShiftDirection.down);
        }
      }
      
      // Write the data
      const dataRange = worksheet.getRangeByIndexes(startRow, startCol, formatted.length, 2);
      dataRange.values = formatted;
      
      // Find and format the "Autor vzorku" row with "Neznámý uživatel"
      for (let i = 0; i < formatted.length; i++) {
        if (formatted[i][0] === 'Autor vzorku' && formatted[i][1] === 'Neznámý uživatel') {
          const authorCell = worksheet.getRangeByIndexes(startRow + i, startCol + 1, 1, 1);
          // Make it bold and yellow background
          authorCell.format.font.bold = true;
          authorCell.format.fill.color = '#FFFF00'; // Yellow background
          authorCell.format.font.color = '#000000'; // Black text for readability
          break;
        }
      }
      
      await context.sync();
    });
  } catch (e) {
    console.error('Chyba při tisku parametrů:', e);
    showNotification('Chyba při tisku parametrů.');
  }
}

// Gathers all sampler parameters from UI and storage
function gatherSamplerParameters() {
  const params = {};
  // Task 2
  params['Kontrolní riziko'] = reliabilityConfig.controlRisk || '';
  params['Přirozené riziko'] = reliabilityConfig.inherentRisk || '';
  params['Hladina RMM'] = reliabilityConfig.rmmLevel || '';
  params['Provádím analytické testy?'] = reliabilityConfig.analyticalTests || '';
  params['Provádím testy kontrol?'] = reliabilityConfig.controlTests || '';
  params['Faktor spolehlivosti (výsledný)'] = _get_value('reliability-factor');
  // Task 3
  params['Typ významnosti'] = _get_value('significance-type-display');
  params['Hodnota významnosti'] = _get_value('significance-value-display');
  params['Zdůvodnění významnosti'] = _get_value('significance-justification-display');
  // Task 4
  params['Oblast dat'] = _get_value('data-range');
  params['Sloupec pro výběr'] = _get_value('value-column');
  // Task 5
  params['Celková suma absolutních hodnot'] = _get_value('task5-result-display');
  params['Vypočtená velikost vzorku'] = _get_value('task5-size-display');
  const useCalculated = _get_value('task5-use-calculated');
  params['Použít vypočtenou velikost'] = useCalculated;
  if (useCalculated === 'Ne') {
    params['Uživatelská velikost vzorku'] = _get_value('task5-override-size');
    params['Důvod uživatelské velikosti'] = _get_value('task5-override-reason');
  }
  // Task 6
  params['Metoda výběru vzorku'] = _get_value('task6-method-display');
  params['Finální velikost vzorku'] = _get_value('task6-final-size-display');
  params['Vyloučit nadlimitní položky'] = _get_value('task6-exclude-over-significance');
  params['Suma pro výpočet (po úpravě)'] = _get_value('task6-param-sum-display');
  params['Počet položek nad významností'] = _get_value('task6-significant-count-display');
  params['Seed pro generátor'] = _get_value('task6-seed');
  // Add new parameters
  params['Autor vzorku'] = "Neznámý uživatel";
  params['Datum a čas vyhotovení'] = new Date().toLocaleString();

  return params;
}
