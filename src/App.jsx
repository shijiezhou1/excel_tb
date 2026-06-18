import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";

const DEFAULT_COL_WIDTH = 140;
const DEFAULT_ROW_HEIGHT = 40;
const MIN_COL_WIDTH = 64;
const MIN_ROW_HEIGHT = 28;
const MAX_COL_WIDTH = 520;
const MAX_ROW_HEIGHT = 260;
const ROW_HEADER_WIDTH = 52;
const COL_HEADER_HEIGHT = 34;
const DEFAULT_FORMULA_HEIGHT = 34;
const MIN_FORMULA_HEIGHT = 34;
const MAX_FORMULA_HEIGHT = 220;
const ROW_BUFFER = 20;

const clamp = (value, min, max) => Math.min(Math.max(value, min), max);

function columnName(index) {
  let value = index + 1;
  let label = "";

  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }

  return label;
}

function columnIndexFromLabel(label) {
  const value = label.trim().toUpperCase();
  if (/^\d+$/.test(value)) return Number(value) - 1;
  if (!/^[A-Z]+$/.test(value)) return null;

  let index = 0;
  for (const letter of value) {
    index = index * 26 + letter.charCodeAt(0) - 64;
  }
  return index - 1;
}

function createDefaultRows(rowCount, colCount) {
  return Array.from({ length: rowCount }, (_, rowIndex) =>
    Array.from({ length: colCount }, (_, colIndex) => {
      if (rowIndex === 0) {
        return columnName(colIndex);
      }
      return "";
    })
  );
}

function normalizeRange(a, b) {
  return {
    startRow: Math.min(a.row, b.row),
    endRow: Math.max(a.row, b.row),
    startCol: Math.min(a.col, b.col),
    endCol: Math.max(a.col, b.col)
  };
}

function isInsideRange(row, col, range) {
  return (
    row >= range.startRow &&
    row <= range.endRow &&
    col >= range.startCol &&
    col <= range.endCol
  );
}

function escapeCellForTsv(value) {
  const text = String(value ?? "");
  if (!/["\n\r\t]/.test(text)) return text;
  return `"${text.replaceAll('"', '""')}"`;
}

function escapeCellForHtml(value) {
  const encoded = String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
  return encoded.replaceAll("\n", "<br>");
}

function rangeToTsv(rows, range) {
  const lines = [];
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    const cells = [];
    for (let col = range.startCol; col <= range.endCol; col += 1) {
      cells.push(escapeCellForTsv(rows[row]?.[col] ?? ""));
    }
    lines.push(cells.join("\t"));
  }
  return lines.join("\n");
}

function rangeToClipboardHtml(rows, range) {
  const parts = ["<table>"];
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    parts.push("<tr>");
    for (let col = range.startCol; col <= range.endCol; col += 1) {
      parts.push(`<td>${escapeCellForHtml(rows[row]?.[col] ?? "")}</td>`);
    }
    parts.push("</tr>");
  }
  parts.push("</table>");
  return parts.join("");
}

function parseClipboardTable(text) {
  const rows = [[]];
  let cell = "";
  let inQuotes = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];

    if (inQuotes) {
      if (char === '"' && next === '"') {
        cell += '"';
        index += 1;
      } else if (char === '"') {
        inQuotes = false;
      } else {
        cell += char;
      }
      continue;
    }

    if (char === '"') {
      inQuotes = true;
    } else if (char === "\t") {
      rows[rows.length - 1].push(cell);
      cell = "";
    } else if (char === "\n") {
      rows[rows.length - 1].push(cell.replace(/\r$/, ""));
      rows.push([]);
      cell = "";
    } else {
      cell += char;
    }
  }

  rows[rows.length - 1].push(cell.replace(/\r$/, ""));

  if (rows.length > 1 && rows[rows.length - 1].length === 1 && rows[rows.length - 1][0] === "") {
    rows.pop();
  }

  return rows;
}

function parseFreezeSpec(spec, maxIndex, parseToken) {
  const indexes = new Set();
  const tokens = spec
    .split(/[,，;；\s]+/)
    .map((token) => token.trim())
    .filter(Boolean);

  tokens.forEach((token) => {
    const rangeParts = token.split(/[:-]/).map((part) => part.trim()).filter(Boolean);
    if (rangeParts.length === 2) {
      const start = parseToken(rangeParts[0]);
      const end = parseToken(rangeParts[1]);
      if (start === null || end === null) return;

      const low = clamp(Math.min(start, end), 0, maxIndex);
      const high = clamp(Math.max(start, end), 0, maxIndex);
      for (let index = low; index <= high; index += 1) {
        indexes.add(index);
      }
      return;
    }

    const index = parseToken(token);
    if (index !== null && index >= 0 && index <= maxIndex) {
      indexes.add(index);
    }
  });

  return Array.from(indexes).sort((a, b) => a - b);
}

function buildPinnedOffsets(indexes, sizes, baseOffset) {
  const offsets = new Map();
  let offset = baseOffset;

  indexes.forEach((index) => {
    offsets.set(index, offset);
    offset += sizes[index];
  });

  return offsets;
}

function sumSizes(sizes, start, end) {
  let total = 0;
  for (let i = start; i < end; i += 1) {
    total += sizes[i] ?? 0;
  }
  return total;
}

function parseRowToken(token) {
  if (!/^\d+$/.test(token)) return null;
  return Number(token) - 1;
}

function parseColumnToken(token) {
  const index = columnIndexFromLabel(token);
  if (index === null || Number.isNaN(index)) {
    return null;
  }
  return index;
}

function generateDemoData(rowCount, colCount) {
  const rows = [];
  for (let r = 0; r < rowCount; r += 1) {
    const row = [];
    for (let c = 0; c < colCount; c += 1) {
      if (r === 0) {
        row.push(columnName(c));
      } else if (c === 0 && r < 8) {
        const labels = [
          "项目A", "任务B", "指标C", "备注D",
          "客户E", "订单F", "库存G", "总计H"
        ];
        row.push(labels[r - 1] || `行${r + 1}`);
      } else if (c < 5) {
        const vals = [
          "已完成", "进行中", "待审核", "高优先级",
          Math.floor(Math.random() * 10000).toString(),
          `值${r}-${c}`
        ];
        row.push(vals[Math.floor(Math.random() * vals.length)]);
      } else if (c === 4 && r < 4) {
        row.push("多行\n文本\n示例\n数据");
      } else {
        row.push(Math.random() > 0.6
          ? Math.floor(Math.random() * 5000).toString()
          : "");
      }
    }
    rows.push(row);
  }
  return rows;
}

function App() {
  const [demoRows, setDemoRows] = useState(() => 100);
  const [demoCols, setDemoCols] = useState(() => 30);
  const [data, setData] = useState(() => generateDemoData(100, 30));

  function applyDemoSize() {
    setData(generateDemoData(demoRows, demoCols));
  }

  return (
    <main className="app-shell">
      <div className="demo-controls">
        <label>
          行数
          <input
            type="number"
            min="1"
            max="2000"
            value={demoRows}
            onChange={(e) => setDemoRows(Math.max(1, Number(e.target.value) || 1))}
          />
        </label>
        <label>
          列数
          <input
            type="number"
            min="1"
            max="5000"
            value={demoCols}
            onChange={(e) => setDemoCols(Math.max(1, Number(e.target.value) || 1))}
          />
        </label>
        <button type="button" onClick={applyDemoSize}>
          生成数据
        </button>
        <span className="demo-info">
          当前 {data.length} 行 × {data[0]?.length ?? 0} 列
        </span>
      </div>
      <section className="sheet-stage" aria-label="Excel style table">
        <ExcelGrid data={data} />
      </section>
    </main>
  );
}

function ExcelGrid({ data }) {
  const rowCount = data.length;
  const colCount = data[0]?.length ?? 0;

  const [rows, setRows] = useState(data);
  const [colWidths, setColWidths] = useState(() =>
    Array.from({ length: colCount }, () => DEFAULT_COL_WIDTH)
  );
  const [rowHeights, setRowHeights] = useState(() =>
    Array.from({ length: rowCount }, () => DEFAULT_ROW_HEIGHT)
  );
  const [formulaHeight, setFormulaHeight] = useState(DEFAULT_FORMULA_HEIGHT);
  const [freezeRowsSpec, setFreezeRowsSpec] = useState("");
  const [freezeColsSpec, setFreezeColsSpec] = useState("");
  const [activeCell, setActiveCell] = useState({ row: 1, col: 0 });
  const [selectionAnchor, setSelectionAnchor] = useState({ row: 1, col: 0 });
  const [selectionFocus, setSelectionFocus] = useState({ row: 1, col: 0 });
  const [editingCell, setEditingCell] = useState(null);
  const [draftValue, setDraftValue] = useState("");
  const [dragSelection, setDragSelection] = useState(null);
  const [fillState, setFillState] = useState(null);
  const [resizeState, setResizeState] = useState(null);
  const [scrollTop, setScrollTop] = useState(0);
  const gridRef = useRef(null);
  const editorRef = useRef(null);

  const selectionRange = useMemo(
    () => normalizeRange(selectionAnchor, selectionFocus),
    [selectionAnchor, selectionFocus]
  );

  const previewFillRange = useMemo(() => {
    if (!fillState?.target) return null;
    const source = fillState.sourceRange;
    return normalizeRange(
      { row: source.startRow, col: source.startCol },
      fillState.target
    );
  }, [fillState]);

  const frozenRows = useMemo(
    () => parseFreezeSpec(freezeRowsSpec, rowCount - 1, parseRowToken),
    [freezeRowsSpec, rowCount]
  );

  const frozenCols = useMemo(
    () => parseFreezeSpec(freezeColsSpec, colCount - 1, parseColumnToken),
    [freezeColsSpec, colCount]
  );

  const frozenRowOffsets = useMemo(
    () => buildPinnedOffsets(frozenRows, rowHeights, COL_HEADER_HEIGHT),
    [frozenRows, rowHeights]
  );

  const frozenColOffsets = useMemo(
    () => buildPinnedOffsets(frozenCols, colWidths, ROW_HEADER_WIDTH),
    [frozenCols, colWidths]
  );

  const frozenRowSet = useMemo(() => new Set(frozenRows), [frozenRows]);
  const frozenColSet = useMemo(() => new Set(frozenCols), [frozenCols]);

  const lastFrozenRow = frozenRows[frozenRows.length - 1] ?? null;
  const lastFrozenCol = frozenCols[frozenCols.length - 1] ?? null;

  const visibleRowRange = useMemo(() => {
    let offset = COL_HEADER_HEIGHT;
    let firstVisible = 0;
    for (let r = 0; r < rowCount; r += 1) {
      const h = rowHeights[r] ?? DEFAULT_ROW_HEIGHT;
      if (offset + h > scrollTop) {
        firstVisible = r;
        break;
      }
      offset += h;
      firstVisible = r + 1;
    }

    const first = Math.max(0, firstVisible - ROW_BUFFER);

    offset = COL_HEADER_HEIGHT;
    const viewportHeight = gridRef.current?.clientHeight ?? 800;
    for (let r = 0; r < first; r += 1) {
      offset += rowHeights[r] ?? DEFAULT_ROW_HEIGHT;
    }
    let last = first;
    for (let r = first; r < rowCount; r += 1) {
      offset += rowHeights[r] ?? DEFAULT_ROW_HEIGHT;
      last = r;
      if (offset > scrollTop + viewportHeight + ROW_BUFFER * DEFAULT_ROW_HEIGHT) break;
    }

    return { first: clamp(first, 0, rowCount - 1), last: clamp(last, 0, rowCount - 1) };
  }, [scrollTop, rowHeights, rowCount]);

  const topSpacerHeight = useMemo(
    () => sumSizes(rowHeights, 0, visibleRowRange.first),
    [rowHeights, visibleRowRange.first]
  );

  const bottomSpacerHeight = useMemo(
    () => sumSizes(rowHeights, visibleRowRange.last + 1, rowCount),
    [rowHeights, visibleRowRange.last, rowCount]
  );

  const spacerRow1 = topSpacerHeight > 0 ? 1 : 0;
  const visibleRowGridStart = 2 + spacerRow1;

  const gridStyle = useMemo(
    () => {
      const cols = [`${ROW_HEADER_WIDTH}px`];
      for (let c = 0; c < colCount; c += 1) {
        cols.push(`${colWidths[c] ?? DEFAULT_COL_WIDTH}px`);
      }

      const rows = [`${COL_HEADER_HEIGHT}px`];
      if (topSpacerHeight > 0) rows.push(`${topSpacerHeight}px`);
      for (let r = visibleRowRange.first; r <= visibleRowRange.last; r += 1) {
        rows.push(`${rowHeights[r] ?? DEFAULT_ROW_HEIGHT}px`);
      }
      if (bottomSpacerHeight > 0) rows.push(`${bottomSpacerHeight}px`);

      return {
        gridTemplateColumns: cols.join(" "),
        gridTemplateRows: rows.join(" ")
      };
    },
    [colWidths, rowHeights, colCount, visibleRowRange, topSpacerHeight, bottomSpacerHeight]
  );

  useEffect(() => {
    if (data !== rows) {
      setRows(data);
    }
  }, [data]);

  useEffect(() => {
    if (colCount !== colWidths.length) {
      setColWidths(Array.from({ length: colCount }, () => DEFAULT_COL_WIDTH));
    }
  }, [colCount]);

  useEffect(() => {
    if (rowCount !== rowHeights.length) {
      setRowHeights(Array.from({ length: rowCount }, () => DEFAULT_ROW_HEIGHT));
    }
  }, [rowCount]);

  useEffect(() => {
    if (editingCell && editorRef.current) {
      editorRef.current.focus();
      editorRef.current.select();
    }
  }, [editingCell]);

  useEffect(() => {
    if (!resizeState) return undefined;

    function handlePointerMove(event) {
      if (resizeState.type === "formula") {
        const nextHeight = clamp(
          resizeState.startSize + event.clientY - resizeState.startPointer,
          MIN_FORMULA_HEIGHT,
          MAX_FORMULA_HEIGHT
        );
        setFormulaHeight(nextHeight);
      } else if (resizeState.type === "col") {
        const nextWidth = clamp(
          resizeState.startSize + event.clientX - resizeState.startPointer,
          MIN_COL_WIDTH,
          MAX_COL_WIDTH
        );
        setColWidths((current) =>
          current.map((width, index) => (index === resizeState.index ? nextWidth : width))
        );
      } else {
        const nextHeight = clamp(
          resizeState.startSize + event.clientY - resizeState.startPointer,
          MIN_ROW_HEIGHT,
          MAX_ROW_HEIGHT
        );
        setRowHeights((current) =>
          current.map((height, index) => (index === resizeState.index ? nextHeight : height))
        );
      }
    }

    function stopResize() {
      setResizeState(null);
    }

    window.addEventListener("pointermove", handlePointerMove);
    window.addEventListener("pointerup", stopResize, { once: true });

    return () => {
      window.removeEventListener("pointermove", handlePointerMove);
      window.removeEventListener("pointerup", stopResize);
    };
  }, [resizeState]);

  useEffect(() => {
    function handlePointerUp() {
      setDragSelection(null);
      setFillState((current) => {
        if (!current?.target) return null;
        applyFill(current.sourceRange, current.target);
        return null;
      });
    }

    window.addEventListener("pointerup", handlePointerUp);
    return () => window.removeEventListener("pointerup", handlePointerUp);
  }, []);

  const commitEdit = useCallback(() => {
    if (!editingCell) return;
    setRows((current) =>
      current.map((row, rowIndex) =>
        rowIndex === editingCell.row
          ? row.map((value, colIndex) => (colIndex === editingCell.col ? draftValue : value))
          : row
      )
    );
    setEditingCell(null);
  }, [draftValue, editingCell]);

  const applyClipboardTable = useCallback((start, table) => {
    setRows((current) =>
      current.map((row, rowIndex) =>
        row.map((cell, colIndex) => {
          const pastedRow = rowIndex - start.row;
          const pastedCol = colIndex - start.col;
          if (
            pastedRow >= 0 &&
            pastedCol >= 0 &&
            pastedRow < table.length &&
            pastedCol < table[pastedRow].length
          ) {
            return table[pastedRow][pastedCol];
          }
          return cell;
        })
      )
    );
  }, []);

  const applyFill = useCallback((sourceRange, target) => {
    const fillRange = normalizeRange(
      { row: sourceRange.startRow, col: sourceRange.startCol },
      target
    );
    const sourceRows = sourceRange.endRow - sourceRange.startRow + 1;
    const sourceCols = sourceRange.endCol - sourceRange.startCol + 1;

    setRows((current) =>
      current.map((row, rowIndex) =>
        row.map((cell, colIndex) => {
          if (!isInsideRange(rowIndex, colIndex, fillRange)) return cell;
          if (isInsideRange(rowIndex, colIndex, sourceRange)) return cell;

          const sourceRow = sourceRange.startRow + ((rowIndex - sourceRange.startRow) % sourceRows);
          const sourceCol = sourceRange.startCol + ((colIndex - sourceRange.startCol) % sourceCols);
          return current[sourceRow][sourceCol];
        })
      )
    );
    setSelectionAnchor({ row: fillRange.startRow, col: fillRange.startCol });
    setSelectionFocus({ row: fillRange.endRow, col: fillRange.endCol });
  }, []);

  const selectCell = useCallback((row, col, extend = false) => {
    const target = { row, col };
    setActiveCell(target);
    if (extend) {
      setSelectionFocus(target);
    } else {
      setSelectionAnchor(target);
      setSelectionFocus(target);
    }
  }, []);

  function selectAllCells() {
    setActiveCell({ row: 0, col: 0 });
    setSelectionAnchor({ row: 0, col: 0 });
    setSelectionFocus({ row: rowCount - 1, col: colCount - 1 });
    gridRef.current?.focus();
  }

  function handleScroll() {
    if (gridRef.current) {
      setScrollTop(gridRef.current.scrollTop);
    }
  }

  function handleFreezeRowsChange(event) {
    setFreezeRowsSpec(event.target.value);
  }

  function handleFreezeColsChange(event) {
    setFreezeColsSpec(event.target.value);
  }

  function getColumnHeaderStyle(col) {
    const style = { gridColumn: col + 2 };
    if (frozenColOffsets.has(col)) {
      style.left = `${frozenColOffsets.get(col)}px`;
    }
    return style;
  }

  function getRowHeaderStyle(row) {
    const gridRow = row >= visibleRowRange.first
      ? visibleRowGridStart + row - visibleRowRange.first
      : 0;
    const style = { gridRow };
    if (frozenRowOffsets.has(row)) {
      style.top = `${frozenRowOffsets.get(row)}px`;
    }
    return style;
  }

  function getCellStyle(row, col) {
    const gridRow = row >= visibleRowRange.first
      ? visibleRowGridStart + row - visibleRowRange.first
      : 0;
    const style = { gridColumn: col + 2, gridRow };
    if (frozenRowOffsets.has(row)) {
      style.top = `${frozenRowOffsets.get(row)}px`;
    }
    if (frozenColOffsets.has(col)) {
      style.left = `${frozenColOffsets.get(col)}px`;
    }
    return style;
  }

  function handleColumnHeaderPointerDown(event, col) {
    if (event.button !== 0 || resizeState) return;
    if (editingCell) commitEdit();
    event.preventDefault();
    const anchorCol = event.shiftKey ? selectionRange.startCol : col;
    setActiveCell({ row: 0, col });
    setSelectionAnchor({ row: 0, col: anchorCol });
    setSelectionFocus({ row: rowCount - 1, col });
    setDragSelection("col");
    gridRef.current?.focus();
  }

  function handleColumnHeaderPointerEnter(col) {
    if (dragSelection !== "col") return;
    setActiveCell({ row: 0, col });
    setSelectionFocus({ row: rowCount - 1, col });
  }

  function handleRowHeaderPointerDown(event, row) {
    if (event.button !== 0 || resizeState) return;
    if (editingCell) commitEdit();
    event.preventDefault();
    const anchorRow = event.shiftKey ? selectionRange.startRow : row;
    setActiveCell({ row, col: 0 });
    setSelectionAnchor({ row: anchorRow, col: 0 });
    setSelectionFocus({ row, col: colCount - 1 });
    setDragSelection("row");
    gridRef.current?.focus();
  }

  function handleRowHeaderPointerEnter(row) {
    if (dragSelection !== "row") return;
    setActiveCell({ row, col: 0 });
    setSelectionFocus({ row, col: colCount - 1 });
  }

  function startColumnResize(event, index) {
    event.preventDefault();
    event.stopPropagation();
    setResizeState({
      type: "col",
      index,
      startPointer: event.clientX,
      startSize: colWidths[index]
    });
  }

  function startRowResize(event, index) {
    event.preventDefault();
    event.stopPropagation();
    setResizeState({
      type: "row",
      index,
      startPointer: event.clientY,
      startSize: rowHeights[index]
    });
  }

  function startFormulaResize(event) {
    event.preventDefault();
    event.stopPropagation();
    setResizeState({
      type: "formula",
      startPointer: event.clientY,
      startSize: formulaHeight
    });
  }

  function startEdit(row, col) {
    setActiveCell({ row, col });
    setSelectionAnchor({ row, col });
    setSelectionFocus({ row, col });
    setDraftValue(rows[row][col]);
    setEditingCell({ row, col });
  }

  function handleCellPointerDown(event, row, col) {
    if (event.button !== 0 || resizeState) return;
    if (editingCell) commitEdit();
    event.preventDefault();
    const extend = event.shiftKey;
    selectCell(row, col, extend);
    setDragSelection("cell");
    gridRef.current?.focus();
  }

  function handleCellPointerEnter(row, col) {
    if (dragSelection === "cell") {
      setSelectionFocus({ row, col });
      setActiveCell({ row, col });
    }

    if (fillState) {
      setFillState((current) => (current ? { ...current, target: { row, col } } : null));
    }
  }

  function handleFillPointerDown(event) {
    event.preventDefault();
    event.stopPropagation();
    setFillState({ sourceRange: selectionRange, target: null });
  }

  function handleCopy(event) {
    if (editingCell) return;
    event.preventDefault();
    event.clipboardData.setData("text/plain", rangeToTsv(rows, selectionRange));
    event.clipboardData.setData("text/html", rangeToClipboardHtml(rows, selectionRange));
  }

  function handleCut(event) {
    if (editingCell) return;
    event.preventDefault();
    event.clipboardData.setData("text/plain", rangeToTsv(rows, selectionRange));
    event.clipboardData.setData("text/html", rangeToClipboardHtml(rows, selectionRange));
    setRows((current) =>
      current.map((row, rowIndex) =>
        row.map((cell, colIndex) =>
          isInsideRange(rowIndex, colIndex, selectionRange) ? "" : cell
        )
      )
    );
  }

  function handlePaste(event) {
    if (editingCell) return;
    const text = event.clipboardData.getData("text/plain");
    if (!text) return;
    event.preventDefault();
    const table = parseClipboardTable(text);
    const start = { row: selectionRange.startRow, col: selectionRange.startCol };
    applyClipboardTable(start, table);
    setSelectionAnchor(start);
    setSelectionFocus({
      row: clamp(start.row + table.length - 1, 0, rowCount - 1),
      col: clamp(start.col + Math.max(...table.map((line) => line.length)) - 1, 0, colCount - 1)
    });
  }

  function handleKeyDown(event) {
    if (editingCell) {
      if (event.key === "Escape") {
        event.preventDefault();
        setEditingCell(null);
      } else if (event.key === "Enter" && (event.metaKey || event.ctrlKey)) {
        event.preventDefault();
        setDraftValue((value) => `${value}\n`);
      } else if (event.key === "Enter") {
        event.preventDefault();
        commitEdit();
        selectCell(clamp(editingCell.row + 1, 0, rowCount - 1), editingCell.col);
      }
      return;
    }

    const next = { ...activeCell };
    if (event.key === "Enter") {
      event.preventDefault();
      startEdit(activeCell.row, activeCell.col);
      return;
    }
    if (event.key === "Backspace" || event.key === "Delete") {
      event.preventDefault();
      setRows((current) =>
        current.map((row, rowIndex) =>
          row.map((cell, colIndex) =>
            isInsideRange(rowIndex, colIndex, selectionRange) ? "" : cell
          )
        )
      );
      return;
    }
    if (event.key.length === 1 && !event.metaKey && !event.ctrlKey && !event.altKey) {
      event.preventDefault();
      setDraftValue(event.key);
      setEditingCell(activeCell);
      return;
    }

    if (event.key === "ArrowUp") next.row = clamp(activeCell.row - 1, 0, rowCount - 1);
    else if (event.key === "ArrowDown") next.row = clamp(activeCell.row + 1, 0, rowCount - 1);
    else if (event.key === "ArrowLeft") next.col = clamp(activeCell.col - 1, 0, colCount - 1);
    else if (event.key === "ArrowRight") next.col = clamp(activeCell.col + 1, 0, colCount - 1);
    else return;

    event.preventDefault();
    selectCell(next.row, next.col, event.shiftKey);
  }

  const activeAddress = `${columnName(activeCell.col)}${activeCell.row + 1}`;
  const isFullRowSelection =
    selectionRange.startCol === 0 && selectionRange.endCol === colCount - 1;
  const isFullColumnSelection =
    selectionRange.startRow === 0 && selectionRange.endRow === rowCount - 1;
  const isAllCellsSelected = isFullRowSelection && isFullColumnSelection;
  const activeRowData = rows[activeCell.row] ?? [];
  const activeCellValue = activeRowData[activeCell.col] ?? "";

  return (
    <div className="workbook" style={{ "--formula-height": `${formulaHeight}px` }}>
      <div className="formula-strip">
        <div className="name-box">{activeAddress}</div>
        <div className="freeze-controls" aria-label="冻结窗格设置">
          <label className="freeze-field">
            <span>冻结行</span>
            <input
              type="text"
              placeholder="1,3 或 1-3"
              value={freezeRowsSpec}
              onChange={handleFreezeRowsChange}
            />
          </label>
          <label className="freeze-field">
            <span>冻结列</span>
            <input
              type="text"
              placeholder="C,D 或 C:D"
              value={freezeColsSpec}
              onChange={handleFreezeColsChange}
            />
          </label>
        </div>
        <div className="formula-box" title={activeCellValue}>
          <div className="formula-content">
            {activeCellValue || "\u00a0"}
          </div>
          <button
            className="formula-resizer"
            aria-label="调整预览栏高度"
            onPointerDown={startFormulaResize}
            type="button"
          />
        </div>
      </div>

      <div
        ref={gridRef}
        className={`grid-viewport ${resizeState ? "is-resizing" : ""}`}
        tabIndex={0}
        onCopy={handleCopy}
        onCut={handleCut}
        onPaste={handlePaste}
        onKeyDown={handleKeyDown}
        onScroll={handleScroll}
        aria-label="Spreadsheet"
      >
        <div className="sheet-grid" style={gridStyle}>
          <div
            className={`corner-cell ${isAllCellsSelected ? "is-selected" : ""}`}
            onPointerDown={selectAllCells}
            role="button"
            tabIndex={-1}
          />

          {Array.from({ length: colCount }, (_, col) => (
            <div
              className={[
                "column-header",
                frozenColOffsets.has(col) ? "is-frozen-col" : "",
                col === lastFrozenCol ? "is-freeze-col-edge" : "",
                isFullColumnSelection &&
                col >= selectionRange.startCol &&
                col <= selectionRange.endCol
                  ? "is-selected"
                  : ""
              ].join(" ")}
              key={`col-${col}`}
              onPointerDown={(event) => handleColumnHeaderPointerDown(event, col)}
              onPointerEnter={() => handleColumnHeaderPointerEnter(col)}
              style={getColumnHeaderStyle(col)}
            >
              <span>{columnName(col)}</span>
              <button
                className="col-resizer"
                aria-label={`调整 ${columnName(col)} 列宽`}
                onPointerDown={(event) => startColumnResize(event, col)}
                type="button"
              />
            </div>
          ))}

          {topSpacerHeight > 0 ? (
            <div className="spacer-cell" style={{ gridColumn: 1, gridRow: 2, height: `${topSpacerHeight}px` }} />
          ) : null}

          {Array.from(
            { length: visibleRowRange.last - visibleRowRange.first + 1 },
            (_, offset) => {
              const row = visibleRowRange.first + offset;
              return (
                <div
                  className={[
                    "row-header",
                    frozenRowOffsets.has(row) ? "is-frozen-row" : "",
                    row === lastFrozenRow ? "is-freeze-row-edge" : "",
                    isFullRowSelection &&
                    row >= selectionRange.startRow &&
                    row <= selectionRange.endRow
                      ? "is-selected"
                      : ""
                  ].join(" ")}
                  key={`row-${row}`}
                  onPointerDown={(event) => handleRowHeaderPointerDown(event, row)}
                  onPointerEnter={() => handleRowHeaderPointerEnter(row)}
                  style={getRowHeaderStyle(row)}
                >
                  <span>{row + 1}</span>
                  <button
                    className="row-resizer"
                    aria-label={`调整第 ${row + 1} 行高`}
                    onPointerDown={(event) => startRowResize(event, row)}
                    type="button"
                  />
                </div>
              );
            }
          )}

          {Array.from(
            { length: visibleRowRange.last - visibleRowRange.first + 1 },
            (_, offset) => {
              const row = visibleRowRange.first + offset;
              const rowData = rows[row] ?? [];

              return Array.from({ length: colCount }, (_, col) => {
                const value = rowData[col] ?? "";
                const selected = isInsideRange(row, col, selectionRange);
                const active = row === activeCell.row && col === activeCell.col;
                const editing = editingCell?.row === row && editingCell?.col === col;
                const frozenRow = frozenRowSet.has(row);
                const frozenCol = frozenColSet.has(col);
                const fillPreview = previewFillRange && isInsideRange(row, col, previewFillRange);
                const isFillHandle =
                  row === selectionRange.endRow &&
                  col === selectionRange.endCol &&
                  !editingCell;

                return (
                  <div
                    className={[
                      "sheet-cell",
                      selected ? "is-selected" : "",
                      active ? "is-active" : "",
                      editing ? "is-editing" : "",
                      frozenRow ? "is-frozen-row" : "",
                      frozenCol ? "is-frozen-col" : "",
                      frozenRow && frozenCol ? "is-frozen-corner" : "",
                      row === lastFrozenRow ? "is-freeze-row-edge" : "",
                      col === lastFrozenCol ? "is-freeze-col-edge" : "",
                      fillPreview ? "is-fill-preview" : ""
                    ].join(" ")}
                    key={`${row}-${col}`}
                    onDoubleClick={() => startEdit(row, col)}
                    onPointerDown={(event) => handleCellPointerDown(event, row, col)}
                    onPointerEnter={() => handleCellPointerEnter(row, col)}
                    role="gridcell"
                    style={getCellStyle(row, col)}
                  >
                    {editing ? (
                      <textarea
                        ref={editorRef}
                        className="cell-editor"
                        value={draftValue}
                        onBlur={commitEdit}
                        onChange={(event) => setDraftValue(event.target.value)}
                        onPointerDown={(event) => event.stopPropagation()}
                      />
                    ) : (
                      <div className="cell-content">{value}</div>
                    )}
                    {isFillHandle ? (
                      <button
                        className="fill-handle"
                        aria-label="拖拽复制选区内容"
                        onPointerDown={handleFillPointerDown}
                        type="button"
                      />
                    ) : null}
                  </div>
                );
              });
            }
          )}

          {bottomSpacerHeight > 0 ? (
            <div className="spacer-cell" style={{ gridColumn: 1, gridRow: visibleRowGridStart + (visibleRowRange.last - visibleRowRange.first + 1), height: `${bottomSpacerHeight}px` }} />
          ) : null}
        </div>
      </div>
    </div>
  );
}

export default App;
