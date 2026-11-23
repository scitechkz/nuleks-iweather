import { useCallback, useMemo, useState } from 'react';
import {
  buildWorkbook, applyCategoryValues, defaultCategories, newCategory, alignBomCsv, isAllowedInput,
  S, extraRatingRow, MAX_EXTRA_RATING_ROWS,
} from '../model/build.js';

const KEY = 'inclement-weather-state-v2';

function load() {
  try {
    const raw = localStorage.getItem(KEY);
    if (raw) return JSON.parse(raw);
  } catch {
    /* storage unavailable */
  }
  return null;
}

function save(state) {
  try {
    localStorage.setItem(KEY, JSON.stringify(state));
  } catch {
    /* storage unavailable or full */
  }
}

const initial = () => ({
  bomGrid: null,
  bomMeta: null,
  categories: defaultCategories(),
  overrides: {},
  extraRatingRows: 0,
  epoch: 0,
});

export function useStore() {
  const [state, setStateRaw] = useState(() => ({ ...initial(), ...(load() || {}) }));
  const setState = useCallback((fn) => {
    setStateRaw((prev) => {
      const next = typeof fn === 'function' ? fn(prev) : fn;
      save(next);
      return next;
    });
  }, []);

  // Structural rebuild only when the CSV, the number of categories, or a reset changes.
  const built = useMemo(
    () => buildWorkbook(state),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [state.bomGrid, state.categories.length, state.extraRatingRows, state.epoch],
  );

  const setOverride = useCallback((sheet, a, raw) => {
    if (!isAllowedInput(built.wb, sheet, a, raw)) return false;
    built.wb.setInput(sheet, a, raw);
    setState((s) => ({ ...s, overrides: { ...s.overrides, [`${sheet}|${a}`]: raw } }));
    return true;
  }, [built, setState]);

  const resetSheet = useCallback((sheet, addrs) => {
    setState((s) => {
      const overrides = {};
      for (const [k, v] of Object.entries(s.overrides)) {
        const [sh, a] = k.split('|');
        if (sh === sheet && (!addrs || addrs.includes(a))) continue;
        overrides[k] = v;
      }
      return { ...s, overrides, epoch: s.epoch + 1 };
    });
  }, [setState]);

  const updateCategory = useCallback((i, patch) => {
    setState((s) => {
      const categories = s.categories.map((c, j) => (j === i ? { ...c, ...patch } : c));
      applyCategoryValues(built.wb, built.layout, categories);
      return { ...s, categories };
    });
  }, [built, setState]);

  const addCategory = useCallback(() => {
    setState((s) => ({ ...s, categories: [...s.categories, newCategory(s.categories.length + 1)] }));
  }, [setState]);

  const removeCategory = useCallback(() => {
    setState((s) => (s.categories.length > 5 ? { ...s, categories: s.categories.slice(0, -1) } : s));
  }, [setState]);

  const resetCategories = useCallback(() => {
    setState((s) => ({ ...s, categories: defaultCategories(), epoch: s.epoch + 1 }));
  }, [setState]);

  // Extra Rain Exposure Rating rows on the Critical Site Works Summary.
  const addRatingRow = useCallback(() => {
    setState((s) => ((s.extraRatingRows || 0) >= MAX_EXTRA_RATING_ROWS ? s
      : { ...s, extraRatingRows: (s.extraRatingRows || 0) + 1 }));
  }, [setState]);

  const removeRatingRow = useCallback(() => {
    setState((s) => {
      const n = s.extraRatingRows || 0;
      if (!n) return s;
      const row = String(extraRatingRow(n - 1));
      const overrides = {};
      for (const [k, v] of Object.entries(s.overrides)) {
        const [sh, a] = k.split('|');
        if (sh === S.CSWS && a.replace(/^[A-Z]+/, '') === row) continue;
        overrides[k] = v;
      }
      return { ...s, overrides, extraRatingRows: n - 1 };
    });
  }, [setState]);

  const loadCsv = useCallback((rows, fileName) => {
    const { grid, warnings, statisticRow } = alignBomCsv(rows);
    setState((s) => ({ ...s, bomGrid: grid, bomMeta: { fileName, warnings, statisticRow, loadedAt: Date.now() } }));
    return warnings;
  }, [setState]);

  // Resets variables, categories and edits; the uploaded csv stays loaded.
  const resetAll = useCallback(() => setState((s) => ({
    ...initial(), bomGrid: s.bomGrid, bomMeta: s.bomMeta, epoch: s.epoch + 1,
  })), [setState]);

  return {
    state, wb: built.wb, layout: built.layout,
    setOverride, resetSheet, updateCategory, addCategory, removeCategory, resetCategories,
    addRatingRow, removeRatingRow,
    loadCsv, resetAll, hasData: Boolean(state.bomGrid),
  };
}
