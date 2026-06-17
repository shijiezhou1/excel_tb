import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";

const ROW_COUNT = 80;
const COL_COUNT = 26;
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

function createInitialRows() {
  return Array.from({ length: ROW_COUNT }, (_, rowIndex) =>
    Array.from({ length: COL_COUNT }, (_, colIndex) => {
      if (rowIndex === 0) {
        const headers = [
          "项目",
          "状态",
          "负责人",
          "备注",
          "长文本示例",
          "下一步"
        ];
        return headers[colIndex] || "";
      }

      if (rowIndex === 1 && colIndex === 0) return "导入数据校验";
      if (rowIndex === 1 && colIndex === 1) return "进行中";
      if (rowIndex === 1 && colIndex === 3) return "支持复制、粘贴、选区拖拽填充。";
      if (rowIndex === 1 && colIndex === 4) {
        return "这个单元格故意放了很长的内容。\n它包含多行文本，行高不会被自动撑开。\n你可以拖拽左侧行边界单独调整这一行，也可以拖拽顶部列边界单独调整这一列。";
      }

      if (rowIndex === 2 && colIndex === 0) return "客户反馈整理";
      if (rowIndex === 2 && colIndex === 4) {
        return "粘贴 Excel/Numbers/表格数据时，会从当前选区左上角开始填入。";
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

function App() {
  return (
    <main className="app-shell">
      <section className="sheet-stage" aria-label="Excel style table">
        <ExcelGrid />
      </section>
    </main>
  );
}

function ExcelGrid() {
  const [rows, setRows] = useState(createInitialRows);
  const [colWidths, setColWidths] = useState(() =>
    Array.from({ length: COL_COUNT }, () => DEFAULT_COL_WIDTH)
  );
  const [rowHeights, setRowHeights] = useState(() =>
    Array.from({ length: ROW_COUNT }, () => DEFAULT_ROW_HEIGHT)
  );
  const [formulaHeight, setFormulaHeight] = useState(DEFAULT_FORMULA_HEIGHT);
  const [activeCell, setActiveCell] = useState({ row: 1, col: 0 });
  const [selectionAnchor, setSelectionAnchor] = useState({ row: 1, col: 0 });
  const [selectionFocus, setSelectionFocus] = useState({ row: 1, col: 0 });
  const [editingCell, setEditingCell] = useState(null);
  const [draftValue, setDraftValue] = useState("");
  const [dragSelection, setDragSelection] = useState(false);
  const [fillState, setFillState] = useState(null);
  const [resizeState, setResizeState] = useState(null);
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

  const gridStyle = useMemo(
    () => ({
      gridTemplateColumns: `${ROW_HEADER_WIDTH}px ${colWidths
        .map((width) => `${width}px`)
        .join(" ")}`,
      gridTemplateRows: `${COL_HEADER_HEIGHT}px ${rowHeights
        .map((height) => `${height}px`)
        .join(" ")}`
    }),
    [colWidths, rowHeights]
  );

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
      setDragSelection(false);
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
    setDragSelection(true);
    gridRef.current?.focus();
  }

  function handleCellPointerEnter(row, col) {
    if (dragSelection) {
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
  }

  function handleCut(event) {
    if (editingCell) return;
    event.preventDefault();
    event.clipboardData.setData("text/plain", rangeToTsv(rows, selectionRange));
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
      row: clamp(start.row + table.length - 1, 0, ROW_COUNT - 1),
      col: clamp(start.col + Math.max(...table.map((line) => line.length)) - 1, 0, COL_COUNT - 1)
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
        selectCell(clamp(editingCell.row + 1, 0, ROW_COUNT - 1), editingCell.col);
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

    if (event.key === "ArrowUp") next.row = clamp(activeCell.row - 1, 0, ROW_COUNT - 1);
    else if (event.key === "ArrowDown") next.row = clamp(activeCell.row + 1, 0, ROW_COUNT - 1);
    else if (event.key === "ArrowLeft") next.col = clamp(activeCell.col - 1, 0, COL_COUNT - 1);
    else if (event.key === "ArrowRight") next.col = clamp(activeCell.col + 1, 0, COL_COUNT - 1);
    else return;

    event.preventDefault();
    selectCell(next.row, next.col, event.shiftKey);
  }

  const activeAddress = `${columnName(activeCell.col)}${activeCell.row + 1}`;

  return (
    <div className="workbook" style={{ "--formula-height": `${formulaHeight}px` }}>
      <div className="formula-strip">
        <div className="name-box">{activeAddress}</div>
        <div className="formula-box" title={rows[activeCell.row][activeCell.col]}>
          <div className="formula-content">
            {rows[activeCell.row][activeCell.col] || "\u00a0"}
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
        aria-label="Spreadsheet"
      >
        <div className="sheet-grid" style={gridStyle}>
          <div className="corner-cell" />

          {Array.from({ length: COL_COUNT }, (_, col) => (
            <div className="column-header" key={`col-${col}`} style={{ gridColumn: col + 2 }}>
              <span>{columnName(col)}</span>
              <button
                className="col-resizer"
                aria-label={`调整 ${columnName(col)} 列宽`}
                onPointerDown={(event) => startColumnResize(event, col)}
                type="button"
              />
            </div>
          ))}

          {Array.from({ length: ROW_COUNT }, (_, row) => (
            <div className="row-header" key={`row-${row}`} style={{ gridRow: row + 2 }}>
              <span>{row + 1}</span>
              <button
                className="row-resizer"
                aria-label={`调整第 ${row + 1} 行高`}
                onPointerDown={(event) => startRowResize(event, row)}
                type="button"
              />
            </div>
          ))}

          {rows.map((rowData, row) =>
            rowData.map((value, col) => {
              const selected = isInsideRange(row, col, selectionRange);
              const active = row === activeCell.row && col === activeCell.col;
              const editing = editingCell?.row === row && editingCell?.col === col;
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
                    fillPreview ? "is-fill-preview" : ""
                  ].join(" ")}
                  key={`${row}-${col}`}
                  onDoubleClick={() => startEdit(row, col)}
                  onPointerDown={(event) => handleCellPointerDown(event, row, col)}
                  onPointerEnter={() => handleCellPointerEnter(row, col)}
                  role="gridcell"
                  style={{ gridColumn: col + 2, gridRow: row + 2 }}
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
            })
          )}
        </div>
      </div>
    </div>
  );
}

export default App;
