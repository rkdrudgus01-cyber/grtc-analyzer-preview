(function () {
  "use strict";

  const GRTC_XLS_MAX_ROWS = 100000;
  const GRTC_XLS_MAX_COLUMNS = 200;
  const GRTC_XLS_MAX_NONEMPTY_CELLS = 1000000;
  const GRTC_XLS_MAX_DENSE_CELLS = 1000000;
  const GRTC_XLS_MAX_DECODED_CHARS = 8 * 1024 * 1024;
  const GRTC_XLS_MAX_CELL_CHARS = 32767;
  const GRTC_XLS_SAFE_ERROR = "파일을 안전하게 읽을 수 없습니다. 파일 형식과 크기를 확인해 주세요.";

  function normalizeCell(value) {
    if (value === undefined || value === null) return "";
    if (value instanceof Date) return value.toISOString();
    return String(value);
  }

  function selectSheet(workbook) {
    for (const name of workbook.SheetNames || []) {
      const sheet = workbook.Sheets[name];
      if (sheet && Object.keys(sheet).some(key => key[0] !== "!" && normalizeCell(window.XLSX.utils.format_cell(sheet[key])).trim().length > 0)) return sheet;
    }
    const firstName = workbook.SheetNames && workbook.SheetNames[0];
    return firstName ? workbook.Sheets[firstName] : null;
  }

  function boundedCell(reference) {
    if (typeof reference !== "string" || !/^[A-Z]{1,3}[1-9][0-9]{0,6}$/.test(reference)) {
      throw new Error("invalid_cell_reference");
    }
    const cell = window.XLSX.utils.decode_cell(reference);
    if (cell.r >= GRTC_XLS_MAX_ROWS || cell.c >= GRTC_XLS_MAX_COLUMNS) {
      throw new Error("coordinate_limit");
    }
    return cell;
  }

  function boundedRange(reference) {
    if (typeof reference !== "string") throw new Error("invalid_range");
    const parts = reference.split(":");
    if (parts.length > 2) throw new Error("invalid_range");
    const start = boundedCell(parts[0]);
    const end = boundedCell(parts[parts.length - 1]);
    if (start.r > end.r || start.c > end.c) throw new Error("reversed_range");
    return { s: start, e: end };
  }

  function enforceWorkbookBounds(workbook) {
    let denseCells = 0;
    let nonEmptyCells = 0;
    let decodedChars = 0;
    for (const name of workbook.SheetNames || []) {
      const sheet = workbook.Sheets[name];
      if (!sheet) throw new Error("missing_sheet");
      const range = sheet["!ref"] === undefined ? null : boundedRange(sheet["!ref"]);
      if (sheet["!fullref"] !== undefined) {
        const full = boundedRange(sheet["!fullref"]);
        if (!range || full.s.r !== range.s.r || full.s.c !== range.s.c || full.e.r !== range.e.r || full.e.c !== range.e.c) {
          throw new Error("truncated_sheet");
        }
      }
      if (range) {
        denseCells += (range.e.r + 1) * (range.e.c + 1);
        if (denseCells > GRTC_XLS_MAX_DENSE_CELLS) throw new Error("dense_cell_limit");
      }
      for (const key of Object.keys(sheet)) {
        if (key[0] === "!") continue;
        const position = boundedCell(key);
        if (!range || position.r < range.s.r || position.c < range.s.c || position.r > range.e.r || position.c > range.e.c) {
          throw new Error("cell_outside_range");
        }
        const cell = sheet[key];
        if (!cell || typeof cell !== "object") throw new Error("invalid_cell");
        if (cell.f !== undefined || cell.F !== undefined) throw new Error("formula_cell_unsupported");
        const characters = Math.max(normalizeCell(cell.v).length, normalizeCell(window.XLSX.utils.format_cell(cell)).length);
        if (characters > GRTC_XLS_MAX_CELL_CHARS) throw new Error("cell_chars_limit");
        if (characters > 0) nonEmptyCells += 1;
        decodedChars += characters;
        if (nonEmptyCells > GRTC_XLS_MAX_NONEMPTY_CELLS) throw new Error("nonempty_cell_limit");
        if (decodedChars > GRTC_XLS_MAX_DECODED_CHARS) throw new Error("decoded_chars_limit");
      }
    }
  }

  function preserveOperationNumbers(sheet) {
    const keys = Object.keys(sheet).filter(key => key[0] !== "!");
    let headerRow = Infinity;
    let headers = [];
    for (const key of keys) {
      const value = normalizeCell(window.XLSX.utils.format_cell(sheet[key])).trim().toUpperCase();
      if (!value) continue;
      const position = window.XLSX.utils.decode_cell(key);
      if (position.r > headerRow) continue;
      if (position.r < headerRow) { headerRow = position.r; headers = []; }
      headers[position.c] = value;
    }
    if (!["TIME", "NEXTSTA", "DIST", "VEL", "ADC"].every(name => headers.some(value => value.includes(name)))) return;
    const exactTime = headers.indexOf("TIME");
    const timeColumn = exactTime !== -1 ? exactTime : headers.findIndex(value => value && value.includes("TIME"));
    for (const key of keys) {
      const cell = sheet[key];
      if (cell.t !== "n" || typeof cell.v !== "number") continue;
      const position = window.XLSX.utils.decode_cell(key);
      if (position.r > headerRow && position.c !== timeColumn) cell.w = String(cell.v);
    }
  }

  window.grtcCanReadLegacyXls = function () {
    return !!(window.XLSX && window.XLSX.read && window.XLSX.utils);
  };

  window.grtcReadXlsRowsFromBase64 = function (base64, reviewLegacyFailure) {
    return new Promise(function (resolve, reject) {
      setTimeout(function () {
        try {
          if (!window.grtcCanReadLegacyXls()) {
            throw new Error("XLS parser is not loaded");
          }
          let workbook;
          try {
            workbook = window.XLSX.read(base64, {
              type: "base64",
              raw: false,
              cellDates: false,
              bookVBA: false,
              cellFormula: true,
              sheetRows: GRTC_XLS_MAX_ROWS + 1,
            });
          } catch (error) {
            // A legacy decoder TypeError yields no rows. Resource and policy
            // failures (including the bounds checks below) remain fatal.
            if (reviewLegacyFailure === true && error instanceof TypeError) {
              resolve("legacy_decode_unreadable");
              return;
            }
            throw error;
          }
          enforceWorkbookBounds(workbook);
          const sheet = selectSheet(workbook);
          if (!sheet) {
            throw new Error("No worksheet found");
          }
          preserveOperationNumbers(sheet);
          const rows = window.XLSX.utils.sheet_to_json(sheet, {
            header: 1,
            raw: false,
            defval: "",
            blankrows: false,
          });
          resolve(rows.filter(function (row) {
            return row.some(function (value) { return normalizeCell(value).trim().length > 0; });
          }).map(function (row) {
            return row.map(normalizeCell);
          }));
        } catch (error) {
          reject(GRTC_XLS_SAFE_ERROR);
        }
      }, 0);
    });
  };
})();
