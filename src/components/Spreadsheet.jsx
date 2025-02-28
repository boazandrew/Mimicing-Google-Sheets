import React, { useState, useRef, useEffect } from "react";
import classNames from "classnames";
import { Bar } from "react-chartjs-2";
import { Chart, registerables } from "chart.js";
import { evaluate } from "mathjs";

Chart.register(...registerables);

const INITIAL_ROWS = 20;
const INITIAL_COLS = 10;
const COLUMN_WIDTH = 100;
const ROW_HEIGHT = 40;

// Helper functions for column/cell references
const columnToIndex = (colLetters) => {
  return (
    colLetters.split("").reduce((acc, letter) => {
      return acc * 26 + (letter.toUpperCase().charCodeAt(0) - 64);
    }, 0) - 1
  );
};

const indexToColumn = (index) => {
  let column = "";
  let temp = index + 1;
  while (temp > 0) {
    temp--;
    column = String.fromCharCode(65 + (temp % 26)) + column;
    temp = Math.floor(temp / 26);
  }
  return column;
};

const cellToIndex = (cellRef) => {
  if (!cellRef || typeof cellRef !== 'string') return [0, 0];
  
  const colMatch = cellRef.match(/[A-Z]+/);
  const rowMatch = cellRef.match(/\d+/);
  
  if (!colMatch || !rowMatch) return [0, 0];
  
  const col = columnToIndex(colMatch[0]);
  const row = parseInt(rowMatch[0]) - 1;
  
  return [row, col];
};

const Spreadsheet = () => {
  const [rows, setRows] = useState(INITIAL_ROWS);
  const [cols, setCols] = useState(INITIAL_COLS);
  const [cells, setCells] = useState(
    Array.from({ length: INITIAL_ROWS }, () => Array(INITIAL_COLS).fill(""))
  );
  const [selectedCell, setSelectedCell] = useState({ row: 0, col: 0 });
  const [selectedRange, setSelectedRange] = useState(null);
  const [formula, setFormula] = useState("");
  const [boldCells, setBoldCells] = useState(new Set());
  const [italicCells, setItalicCells] = useState(new Set());
  const [fontSize, setFontSize] = useState(14);
  const [color, setColor] = useState("#000000");
  const [columnWidths, setColumnWidths] = useState(
    Array(INITIAL_COLS).fill(COLUMN_WIDTH)
  );
  const [rowHeights, setRowHeights] = useState(
    Array(INITIAL_ROWS).fill(ROW_HEIGHT)
  );
  const [validations, setValidations] = useState(
    Array.from({ length: INITIAL_ROWS }, () => Array(INITIAL_COLS).fill("any"))
  );
  const [errors, setErrors] = useState(new Set());
  const [findReplace, setFindReplace] = useState({
    show: false,
    find: "",
    replace: "",
    currentMatch: 0,
    matches: [],
  });
  const [chartData, setChartData] = useState(null);
  const [dependencies, setDependencies] = useState({});
  const containerRef = useRef(null);
  const [evaluatedValues, setEvaluatedValues] = useState(
    Array.from({ length: INITIAL_ROWS }, () => Array(INITIAL_COLS).fill(""))
  );
  const [sortConfig, setSortConfig] = useState({
    show: false,
    column: 0,
    direction: 'asc'
  });

  // Initialize cells with evaluated values
  useEffect(() => {
    updateAllEvaluatedValues();
  }, []);

  // Update all evaluated values
  const updateAllEvaluatedValues = () => {
    // First pass: collect all dependencies
    const deps = {};
    cells.forEach((rowCells, rIdx) => {
      rowCells.forEach((cell, cIdx) => {
        if (typeof cell === 'string' && cell.startsWith('=')) {
          const cellKey = `${rIdx}-${cIdx}`;
          const cellDeps = getCellDependencies(cell);
          deps[cellKey] = cellDeps;
        }
      });
    });
    
    setDependencies(deps);
    
    // Second pass: evaluate all cells in dependency order
    const newEvaluatedValues = Array.from({ length: rows }, () => Array(cols).fill(""));
    const evaluated = new Set();
    
    // Helper function to evaluate a cell with its dependencies
    const evaluateCell = (row, col, visited = new Set()) => {
      const cellKey = `${row}-${col}`;
      
      // If already evaluated, return the value
      if (evaluated.has(cellKey)) {
        return newEvaluatedValues[row][col];
      }
      
      // Check for circular dependencies
      if (visited.has(cellKey)) {
        return "#CIRCULAR!";
      }
      
      visited.add(cellKey);
      
      const cellValue = cells[row][col];
      
      // If not a formula, return as is
      if (typeof cellValue !== 'string' || !cellValue.startsWith('=')) {
        newEvaluatedValues[row][col] = cellValue;
        evaluated.add(cellKey);
        return cellValue;
      }
      
      // Evaluate dependencies first
      const cellDeps = deps[cellKey] || [];
      for (const dep of cellDeps) {
        const [depRow, depCol] = dep.split('-').map(Number);
        if (depRow >= 0 && depRow < rows && depCol >= 0 && depCol < cols) {
          evaluateCell(depRow, depCol, new Set([...visited]));
        }
      }
      
      // Now evaluate this cell
      try {
        const result = evaluateFormula(cellValue, row, col, newEvaluatedValues);
        newEvaluatedValues[row][col] = result;
        evaluated.add(cellKey);
        return result;
      } catch (error) {
        newEvaluatedValues[row][col] = "#ERROR!";
        evaluated.add(cellKey);
        return "#ERROR!";
      }
    };
    
    // Evaluate all cells
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        evaluateCell(r, c);
      }
    }
    
    setEvaluatedValues(newEvaluatedValues);
  };

  // Get cell dependencies from a formula
  const getCellDependencies = (formula) => {
    if (!formula || typeof formula !== 'string' || !formula.startsWith('=')) {
      return [];
    }
    
    const deps = new Set();
    const cellRefs = formula.match(/[A-Z]+\d+/g) || [];
    
    for (const ref of cellRefs) {
      const [row, col] = cellToIndex(ref);
      if (row >= 0 && row < rows && col >= 0 && col < cols) {
        deps.add(`${row}-${col}`);
      }
    }
    
    return Array.from(deps);
  };

  // Formula evaluation
  const evaluateFormula = (formula, row, col, evalValues) => {
    if (!formula || typeof formula !== 'string') return "";
    if (!formula.startsWith("=")) return formula;
    
    try {
      let expression = formula.substring(1).trim();
      
      // Handle built-in functions
      if (expression.toUpperCase().startsWith("SUM(")) {
        const rangeStr = expression.substring(4, expression.length - 1);
        return calculateSum(rangeStr, evalValues);
      }
      
      if (expression.toUpperCase().startsWith("AVERAGE(")) {
        const rangeStr = expression.substring(8, expression.length - 1);
        return calculateAverage(rangeStr, evalValues);
      }
      
      if (expression.toUpperCase().startsWith("MAX(")) {
        const rangeStr = expression.substring(4, expression.length - 1);
        return calculateMax(rangeStr, evalValues);
      }
      
      if (expression.toUpperCase().startsWith("MIN(")) {
        const rangeStr = expression.substring(4, expression.length - 1);
        return calculateMin(rangeStr, evalValues);
      }
      
      if (expression.toUpperCase().startsWith("COUNT(")) {
        const rangeStr = expression.substring(6, expression.length - 1);
        return calculateCount(rangeStr, evalValues);
      }
      
      // Replace cell references with actual values
      expression = expression.replace(/[A-Z]+\d+/g, (match) => {
        const [refRow, refCol] = cellToIndex(match);
        if (refRow >= 0 && refRow < rows && refCol >= 0 && refCol < cols) {
          const cellValue = evalValues ? evalValues[refRow][refCol] : evaluatedValues[refRow][refCol];
          
          // If the value is a string that's not a number, wrap it in quotes
          if (typeof cellValue === 'string' && isNaN(parseFloat(cellValue))) {
            return `"${cellValue.replace(/"/g, '\\"')}"`;
          }
          
          return cellValue === "" ? 0 : cellValue;
        }
        return 0;
      });
      
      // Safely evaluate the expression
      const result = evaluate(expression);
      return typeof result === 'number' ? parseFloat(result.toFixed(10)) : result;
    } catch (error) {
      console.error("Formula evaluation error:", error, formula);
      return "#ERROR!";
    }
  };
  
  // Range calculation functions
  const parseRange = (rangeStr) => {
    const parts = rangeStr.split(':');
    if (parts.length === 1) {
      // Single cell reference
      return [cellToIndex(parts[0])];
    } else if (parts.length === 2) {
      // Range reference (e.g., A1:B3)
      const [startRow, startCol] = cellToIndex(parts[0]);
      const [endRow, endCol] = cellToIndex(parts[1]);
      
      const cells = [];
      for (let r = Math.min(startRow, endRow); r <= Math.max(startRow, endRow); r++) {
        for (let c = Math.min(startCol, endCol); c <= Math.max(startCol, endCol); c++) {
          cells.push([r, c]);
        }
      }
      return cells;
    }
    return [];
  };
  
  const calculateSum = (rangeStr, evalValues) => {
    const cellIndices = parseRange(rangeStr);
    let sum = 0;
    
    cellIndices.forEach(([r, c]) => {
      if (r >= 0 && r < rows && c >= 0 && c < cols) {
        const cellValue = evalValues ? evalValues[r][c] : evaluatedValues[r][c];
        const numValue = parseFloat(cellValue);
        if (!isNaN(numValue)) {
          sum += numValue;
        }
      }
    });
    
    return sum;
  };
  
  const calculateAverage = (rangeStr, evalValues) => {
    const cellIndices = parseRange(rangeStr);
    let sum = 0;
    let count = 0;
    
    cellIndices.forEach(([r, c]) => {
      if (r >= 0 && r < rows && c >= 0 && c < cols) {
        const cellValue = evalValues ? evalValues[r][c] : evaluatedValues[r][c];
        const numValue = parseFloat(cellValue);
        if (!isNaN(numValue)) {
          sum += numValue;
          count++;
        }
      }
    });
    
    return count > 0 ? sum / count : 0;
  };
  
  const calculateMax = (rangeStr, evalValues) => {
    const cellIndices = parseRange(rangeStr);
    let max = Number.NEGATIVE_INFINITY;
    let hasValue = false;
    
    cellIndices.forEach(([r, c]) => {
      if (r >= 0 && r < rows && c >= 0 && c < cols) {
        const cellValue = evalValues ? evalValues[r][c] : evaluatedValues[r][c];
        const numValue = parseFloat(cellValue);
        if (!isNaN(numValue)) {
          max = Math.max(max, numValue);
          hasValue = true;
        }
      }
    });
    
    return hasValue ? max : 0;
  };
  
  const calculateMin = (rangeStr, evalValues) => {
    const cellIndices = parseRange(rangeStr);
    let min = Number.POSITIVE_INFINITY;
    let hasValue = false;
    
    cellIndices.forEach(([r, c]) => {
      if (r >= 0 && r < rows && c >= 0 && c < cols) {
        const cellValue = evalValues ? evalValues[r][c] : evaluatedValues[r][c];
        const numValue = parseFloat(cellValue);
        if (!isNaN(numValue)) {
          min = Math.min(min, numValue);
          hasValue = true;
        }
      }
    });
    
    return hasValue ? min : 0;
  };
  
  const calculateCount = (rangeStr, evalValues) => {
    const cellIndices = parseRange(rangeStr);
    let count = 0;
    
    cellIndices.forEach(([r, c]) => {
      if (r >= 0 && r < rows && c >= 0 && c < cols) {
        const cellValue = evalValues ? evalValues[r][c] : evaluatedValues[r][c];
        const numValue = parseFloat(cellValue);
        if (!isNaN(numValue)) {
          count++;
        }
      }
    });
    
    return count;
  };

  const handleCellChange = (row, col, value) => {
    const newCells = [...cells];
    newCells[row][col] = value;
    setCells(newCells);
    
    // Update the formula bar when a cell is selected
    if (row === selectedCell.row && col === selectedCell.col) {
      setFormula(value);
    }
    
    // Validate the cell value based on its validation type
    validateCell(row, col, value);
    
    // Update dependencies and evaluated values
    updateAllEvaluatedValues();
  };
  
  // Validate cell based on validation type
  const validateCell = (row, col, value) => {
    const validationType = validations[row][col];
    const cellKey = `${row}-${col}`;
    const newErrors = new Set(errors);
    
    if (value === "" || value.startsWith("=")) {
      newErrors.delete(cellKey);
    } else if (validationType === "number") {
      const numValue = parseFloat(value);
      if (isNaN(numValue)) {
        newErrors.add(cellKey);
      } else {
        newErrors.delete(cellKey);
      }
    } else if (validationType === "date") {
      const dateValue = new Date(value);
      if (isNaN(dateValue.getTime())) {
        newErrors.add(cellKey);
      } else {
        newErrors.delete(cellKey);
      }
    } else {
      newErrors.delete(cellKey);
    }
    
    setErrors(newErrors);
  };

  // Data Quality Functions
  const applyToRange = (transformFunc) => {
    // If no range is selected, apply to the current cell
    if (!selectedRange) {
      const { row, col } = selectedCell;
      const newCells = [...cells];
      if (typeof newCells[row][col] === "string") {
        newCells[row][col] = transformFunc(newCells[row][col]);
        setCells(newCells);
        updateAllEvaluatedValues();
      }
      return;
    }
    
    // If range is selected, apply to all cells in range
    const newCells = [...cells];
    
    for (let r = selectedRange.startRow; r <= selectedRange.endRow; r++) {
      for (let c = selectedRange.startCol; c <= selectedRange.endCol; c++) {
        if (r >= 0 && r < rows && c >= 0 && c < cols) {
          const cellValue = newCells[r][c];
          if (typeof cellValue === "string") {
            newCells[r][c] = transformFunc(cellValue);
          }
        }
      }
    }
    
    setCells(newCells);
    updateAllEvaluatedValues();
  };

  const removeDuplicates = () => {
    // If no range is selected, use all cells
    if (!selectedRange) {
      // Create a map to track unique values
      const uniqueValues = new Map();
      const newCells = [...cells];
      
      // First pass: identify unique values
      for (let r = 0; r < rows; r++) {
        const rowKey = newCells[r].join('|');
        if (!uniqueValues.has(rowKey)) {
          uniqueValues.set(rowKey, r);
        }
      }
      
      // Second pass: keep only unique rows
      const uniqueRows = Array.from(uniqueValues.values()).sort((a, b) => a - b);
      const filteredCells = uniqueRows.map(r => [...newCells[r]]);
      
      setCells(filteredCells);
      setRows(filteredCells.length);
      updateAllEvaluatedValues();
      return;
    }
    
    // If range is selected, remove duplicates within that range
    const seen = new Set();
    const newCells = [...cells];
    const rowsToRemove = new Set();
    
    // Identify duplicate rows
    for (let r = selectedRange.startRow; r <= selectedRange.endRow; r++) {
      const rowKey = newCells[r].slice(selectedRange.startCol, selectedRange.endCol + 1).join('|');
      if (seen.has(rowKey)) {
        rowsToRemove.add(r);
      } else {
        seen.add(rowKey);
      }
    }
    
    // Remove rows in reverse order to avoid index shifting
    const rowsToRemoveArray = Array.from(rowsToRemove).sort((a, b) => b - a);
    for (const r of rowsToRemoveArray) {
      newCells.splice(r, 1);
    }
    
    setCells(newCells);
    setRows(newCells.length);
    updateAllEvaluatedValues();
  };

  // Sort functionality
  const sortData = () => {
    if (!selectedRange) {
      // Sort the entire spreadsheet based on the selected column
      const { col } = selectedCell;
      const newCells = [...cells];
      
      newCells.sort((rowA, rowB) => {
        const valueA = parseFloat(evaluatedValues[newCells.indexOf(rowA)][col]);
        const valueB = parseFloat(evaluatedValues[newCells.indexOf(rowB)][col]);
        
        // Handle non-numeric values
        const isNumA = !isNaN(valueA);
        const isNumB = !isNaN(valueB);
        
        if (!isNumA && !isNumB) {
          // Both are strings, compare alphabetically
          return sortConfig.direction === 'asc' 
            ? String(rowA[col]).localeCompare(String(rowB[col]))
            : String(rowB[col]).localeCompare(String(rowA[col]));
        } else if (!isNumA) {
          // A is string, B is number, strings come after numbers in ascending order
          return sortConfig.direction === 'asc' ? 1 : -1;
        } else if (!isNumB) {
          // B is string, A is number, strings come after numbers in ascending order
          return sortConfig.direction === 'asc' ? -1 : 1;
        } else {
          // Both are numbers, compare numerically
          return sortConfig.direction === 'asc' ? valueA - valueB : valueB - valueA;
        }
      });
      
      setCells(newCells);
    } else {
      // Sort only the selected range
      const { startRow, endRow, startCol, endCol } = selectedRange;
      const sortCol = sortConfig.column - startCol;
      
      if (sortCol < 0 || sortCol > (endCol - startCol)) {
        alert("Sort column must be within the selected range");
        return;
      }
      
      const newCells = [...cells];
      const rangeCells = [];
      
      // Extract the range
      for (let r = startRow; r <= endRow; r++) {
        rangeCells.push(newCells[r].slice(startCol, endCol + 1));
      }
      
      // Sort the range
      rangeCells.sort((rowA, rowB) => {
        const valueA = parseFloat(rowA[sortCol]);
        const valueB = parseFloat(rowB[sortCol]);
        
        // Handle non-numeric values
        const isNumA = !isNaN(valueA);
        const isNumB = !isNaN(valueB);
        
        if (!isNumA && !isNumB) {
          return sortConfig.direction === 'asc' 
            ? String(rowA[sortCol]).localeCompare(String(rowB[sortCol]))
            : String(rowB[sortCol]).localeCompare(String(rowA[sortCol]));
        } else if (!isNumA) {
          return sortConfig.direction === 'asc' ? 1 : -1;
        } else if (!isNumB) {
          return sortConfig.direction === 'asc' ? -1 : 1;
        } else {
          return sortConfig.direction === 'asc' ? valueA - valueB : valueB - valueA;
        }
      });
      
      // Put the sorted range back
      for (let r = 0; r < rangeCells.length; r++) {
        for (let c = 0; c < rangeCells[r].length; c++) {
          newCells[startRow + r][startCol + c] = rangeCells[r][c];
        }
      }
      
      setCells(newCells);
    }
    
    updateAllEvaluatedValues();
    setSortConfig({ ...sortConfig, show: false });
  };

  // Find/Replace functionality
  const findMatches = () => {
    if (!findReplace.find) return;
    
    const matches = [];
    cells.forEach((row, rowIndex) => {
      row.forEach((cell, colIndex) => {
        if (typeof cell === "string" && cell.includes(findReplace.find)) {
          matches.push({ row: rowIndex, col: colIndex });
        }
      });
    });
    
    setFindReplace((prev) => ({
      ...prev,
      matches,
      currentMatch: matches.length > 0 ? 0 : -1,
    }));
    
    // Select the first match if found
    if (matches.length > 0) {
      setSelectedCell({ row: matches[0].row, col: matches[0].col });
    }
  };

  const handleReplace = () => {
    const { matches, currentMatch, find, replace } = findReplace;
    if (matches.length === 0 || currentMatch < 0) return;

    const newCells = [...cells];
    const { row, col } = matches[currentMatch];
    newCells[row][col] = newCells[row][col].replace(find, replace);

    setCells(newCells);
    updateAllEvaluatedValues();

    // Move to next match or reset
    if (currentMatch < matches.length - 1) {
      const nextMatch = currentMatch + 1;
      setFindReplace((prev) => ({
        ...prev,
        currentMatch: nextMatch,
      }));
      setSelectedCell({ row: matches[nextMatch].row, col: matches[nextMatch].col });
    } else {
      setFindReplace((prev) => ({ ...prev, matches: [], currentMatch: -1 }));
    }
  };

  const handleReplaceAll = () => {
    const { find, replace } = findReplace;
    if (!find) return;
    
    const newCells = cells.map((row) =>
      row.map((cell) =>
        typeof cell === "string" ? cell.replace(new RegExp(find, 'g'), replace) : cell
      )
    );
    
    setCells(newCells);
    updateAllEvaluatedValues();
    setFindReplace((prev) => ({ ...prev, matches: [], currentMatch: -1 }));
  };

  // Save/Load functionality
  const saveSpreadsheet = () => {
    const data = JSON.stringify({
      cells,
      columnWidths,
      rowHeights,
      validations,
      boldCells: Array.from(boldCells),
      italicCells: Array.from(italicCells),
      fontSize,
      color,
    });
    localStorage.setItem("spreadsheet", data);
    alert("Spreadsheet saved successfully!");
  };

  const loadSpreadsheet = () => {
    try {
      const savedData = localStorage.getItem("spreadsheet");
      if (!savedData) {
        alert("No saved spreadsheet found!");
        return;
      }
      
      const data = JSON.parse(savedData);
      
      setCells(data.cells || []);
      setColumnWidths(data.columnWidths || []);
      setRowHeights(data.rowHeights || []);
      setValidations(data.validations || []);
      setBoldCells(new Set(data.boldCells || []));
      setItalicCells(new Set(data.italicCells || []));
      setFontSize(data.fontSize || 14);
      setColor(data.color || "#000000");
      
      // Update rows and columns count
      if (data.cells) {
        setRows(data.cells.length);
        setCols(data.cells[0]?.length || INITIAL_COLS);
      }
      
      updateAllEvaluatedValues();
      alert("Spreadsheet loaded successfully!");
    } catch (error) {
      console.error("Error loading spreadsheet:", error);
      alert("Error loading spreadsheet!");
    }
  };

  // Data Visualization
  const createChart = () => {
    // If no range is selected, use the current cell
    if (!selectedRange) {
      const { row, col } = selectedCell;
      const cellValue = evaluatedValues[row][col];
      const value = parseFloat(cellValue);
      
      setChartData({
        labels: [`${indexToColumn(col)}${row + 1}`],
        datasets: [
          {
            label: "Cell Value",
            data: [isNaN(value) ? 0 : value],
            backgroundColor: "rgba(54, 162, 235, 0.2)",
            borderColor: "rgba(54, 162, 235, 1)",
            borderWidth: 1,
          },
        ],
      });
      return;
    }

    // If range is selected, use all cells in the range
    const labels = [];
    const dataPoints = [];

    for (let r = selectedRange.startRow; r <= selectedRange.endRow; r++) {
      for (let c = selectedRange.startCol; c <= selectedRange.endCol; c++) {
        if (r >= 0 && r < rows && c >= 0 && c < cols) {
          labels.push(`${indexToColumn(c)}${r + 1}`);
          const cellValue = evaluatedValues[r][c];
          const value = parseFloat(cellValue);
          dataPoints.push(isNaN(value) ? 0 : value);
        }
      }
    }

    setChartData({
      labels,
      datasets: [
        {
          label: "Cell Values",
          data: dataPoints,
          backgroundColor: "rgba(54, 162, 235, 0.2)",
          borderColor: "rgba(54, 162, 235, 1)",
          borderWidth: 1,
        },
      ],
    });
  };
  
  // Handle cell selection with shift key for range selection
  const handleCellClick = (row, col, isShiftKey) => {
    if (isShiftKey && selectedCell) {
      // Create a range selection
      setSelectedRange({
        startRow: Math.min(selectedCell.row, row),
        startCol: Math.min(selectedCell.col, col),
        endRow: Math.max(selectedCell.row, row),
        endCol: Math.max(selectedCell.col, col),
      });
    } else {
      // Single cell selection
      setSelectedCell({ row, col });
      setSelectedRange(null);
      setFormula(cells[row][col] || "");
    }
  };
  
  // Drag and Drop handlers
  const handleDragStart = (e, row, col) => {
    e.dataTransfer.setData("text/plain", JSON.stringify({ 
      value: cells[row][col],
      sourceRow: row,
      sourceCol: col
    }));
  };

  const handleDragOver = (e) => {
    e.preventDefault();
  };

  const handleDrop = (e, row, col) => {
    e.preventDefault();
    try {
      const data = JSON.parse(e.dataTransfer.getData("text/plain"));
      const { value, sourceRow, sourceCol } = data;
      
      // If shift key is pressed, copy the value
      if (e.shiftKey) {
        handleCellChange(row, col, value);
      } else {
        // Move the value (clear source cell and set target cell)
        const newCells = [...cells];
        newCells[sourceRow][sourceCol] = "";
        newCells[row][col] = value;
        setCells(newCells);
        updateAllEvaluatedValues();
      }
    } catch (error) {
      console.error("Error during drag and drop:", error);
    }
  };
  
  // Row and column management
  const addRow = () => {
    setRows((prev) => prev + 1);
    setCells((prev) => [...prev, Array(cols).fill("")]);
    setRowHeights((prev) => [...prev, ROW_HEIGHT]);
    setEvaluatedValues((prev) => [...prev, Array(cols).fill("")]);
    setValidations((prev) => [...prev, Array(cols).fill("any")]);
  };

  const deleteRow = () => {
    if (rows > 1) {
      setRows((prev) => prev - 1);
      setCells((prev) => prev.slice(0, -1));
      setRowHeights((prev) => prev.slice(0, -1));
      setEvaluatedValues((prev) => prev.slice(0, -1));
      setValidations((prev) => prev.slice(0, -1));
    }
  };

  const addColumn = () => {
    setCols((prev) => prev + 1);
    setCells((prev) => prev.map((row) => [...row, ""]));
    setColumnWidths((prev) => [...prev, COLUMN_WIDTH]);
    setEvaluatedValues((prev) => prev.map((row) => [...row, ""]));
    setValidations((prev) => prev.map((row) => [...row, "any"]));
  };

  const deleteColumn = () => {
    if (cols > 1) {
      setCols((prev) => prev - 1);
      setCells((prev) => prev.map((row) => row.slice(0, -1)));
      setColumnWidths((prev) => prev.slice(0, -1));
      setEvaluatedValues((prev) => prev.map((row) => row.slice(0, -1)));
      setValidations((prev) => prev.map((row) => row.slice(0, -1)));
    }
  };
  
  // Resize handlers
  const handleColumnResize = (col, e) => {
    const startX = e.clientX;
    const startWidth = columnWidths[col];

    const doDrag = (moveEvent) => {
      const newWidth = Math.max(40, startWidth + (moveEvent.clientX - startX));
      setColumnWidths((prev) => {
        const newWidths = [...prev];
        newWidths[col] = newWidth;
        return newWidths;
      });
    };

    const stopDrag = () => {
      document.removeEventListener("mousemove", doDrag);
      document.removeEventListener("mouseup", stopDrag);
    };

    document.addEventListener("mousemove", doDrag);
    document.addEventListener("mouseup", stopDrag);
  };

  const handleRowResize = (row, e) => {
    const startY = e.clientY;
    const startHeight = rowHeights[row];

    const doDrag = (moveEvent) => {
      const newHeight = Math.max(
        20,
        startHeight + (moveEvent.clientY - startY)
      );
      setRowHeights((prev) => {
        const newHeights = [...prev];
        newHeights[row] = newHeight;
        return newHeights;
      });
    };

    const stopDrag = () => {
      document.removeEventListener("mousemove", doDrag);
      document.removeEventListener("mouseup", stopDrag);
    };

    document.addEventListener("mousemove", doDrag);
    document.addEventListener("mouseup", stopDrag);
  };

  return (
    <div className="p-4" ref={containerRef}>
      {/* Toolbar */}
      <div className="flex flex-wrap gap-2 mb-4 p-2 bg-gray-100 rounded">
        <button
          onClick={addRow}
          className="px-2 py-1 bg-blue-500 text-white rounded hover:bg-blue-600"
        >
          Add Row
        </button>
        <button
          onClick={deleteRow}
          className="px-2 py-1 bg-red-500 text-white rounded hover:bg-red-600"
        >
          Delete Row
        </button>
        <button
          onClick={addColumn}
          className="px-2 py-1 bg-blue-500 text-white rounded hover:bg-blue-600"
        >
          Add Column
        </button>
        <button
          onClick={deleteColumn}
          className="px-2 py-1 bg-red-500 text-white rounded hover:bg-red-600"
        >
          Delete Column
        </button>

        <div className="h-6 border-l border-gray-300 mx-1"></div>

        <button
          onClick={() => {
            const cellKey = `${selectedCell.row}-${selectedCell.col}`;
            setBoldCells((prev) => {
              const newSet = new Set(prev);
              newSet.has(cellKey)
                ? newSet.delete(cellKey)
                : newSet.add(cellKey);
              return newSet;
            });
          }}
          className="px-2 py-1 bg-gray-300 hover:bg-gray-400 rounded font-bold"
        >
          B
        </button>

        <button
          onClick={() => {
            const cellKey = `${selectedCell.row}-${selectedCell.col}`;
            setItalicCells((prev) => {
              const newSet = new Set(prev);
              newSet.has(cellKey)
                ? newSet.delete(cellKey)
                : newSet.add(cellKey);
              return newSet;
            });
          }}
          className="px-2 py-1 bg-gray-300 hover:bg-gray-400 rounded italic"
        >
          I
        </button>

        <input
          type="number"
          value={fontSize}
          onChange={(e) => setFontSize(parseInt(e.target.value) || 14)}
          className="w-16 px-2 py-1 border rounded"
          min="8"
          max="36"
        />

        <input
          type="color"
          value={color}
          onChange={(e) => setColor(e.target.value)}
          className="w-10 h-10"
        />

        <div className="h-6 border-l border-gray-300 mx-1"></div>

        {/* Data Quality Functions */}
        <button
          onClick={() => applyToRange((value) => value.trim())}
          className="px-2 py-1 bg-green-500 text-white rounded hover:bg-green-600"
        >
          TRIM
        </button>
        <button
          onClick={() => applyToRange((value) => value.toUpperCase())}
          className="px-2 py-1 bg-blue-500 text-white rounded hover:bg-blue-600"
        >
          UPPER
        </button>
        <button
          onClick={() => applyToRange((value) => value.toLowerCase())}
          className="px-2 py-1 bg-yellow-500 text-white rounded hover:bg-yellow-600"
        >
          LOWER
        </button>
        <button
          onClick={removeDuplicates}
          className="px-2 py-1 bg-green-500 text-white rounded hover:bg-green-600"
        >
          Remove Dups
        </button>
        <button
          onClick={() => setFindReplace((p) => ({ ...p, show: true }))}
          className="px-2 py-1 bg-green-500 text-white rounded hover:bg-green-600"
        >
          Find/Replace
        </button>
        <button
          onClick={() => setSortConfig((p) => ({ ...p, show: true }))}
          className="px-2 py-1 bg-purple-500 text-white rounded hover:bg-purple-600"
        >
          Sort
        </button>

        <div className="h-6 border-l border-gray-300 mx-1"></div>

        {/* Data Validation */}
        <select
          value={validations[selectedCell.row]?.[selectedCell.col] || "any"}
          onChange={(e) => {
            const newValidations = [...validations];
            newValidations[selectedCell.row][selectedCell.col] = e.target.value;
            setValidations(newValidations);
          }}
          className="px-2 py-1 border rounded"
        >
          <option value="any">Any</option>
          <option value="number">Number</option>
          <option value="date">Date</option>
        </select>

        <div className="h-6 border-l border-gray-300 mx-1"></div>

        {/* Save/Load */}
        <button
          onClick={saveSpreadsheet}
          className="px-2 py-1 bg-purple-500 text-white rounded hover:bg-purple-600"
        >
          Save
        </button>
        <button
          onClick={loadSpreadsheet}
          className="px-2 py-1 bg-purple-500 text-white rounded hover:bg-purple-600"
        >
          Load
        </button>

        {/* Chart */}
        <button
          onClick={createChart}
          className="px-2 py-1 bg-yellow-500 text-white rounded hover:bg-yellow-600"
        >
          Chart
        </button>
      </div>

      {/* Formula Bar */}
      <div className="mb-4">
        <div className="flex items-center gap-2">
          <div className="bg-gray-100 px-2 py-1 border rounded">
            {selectedCell ? `${indexToColumn(selectedCell.col)}${selectedCell.row + 1}` : ""}
          </div>
          <input
            type="text"
            value={formula}
            onChange={(e) => setFormula(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter") {
                const value = formula.trim();
                if (selectedCell) {
                  handleCellChange(selectedCell.row, selectedCell.col, value);
                }
              }
            }}
            className="w-full p-2 border rounded"
            placeholder="Enter value or formula (e.g., =SUM(A1:A5))"
          />
        </div>
      </div>

      {/* Spreadsheet Grid */}
      <div className="border rounded overflow-auto">
        {/* Column Headers */}
        <div className="flex">
          <div className="w-12 bg-gray-100 border-r border-b"></div>
          {Array.from({ length: cols }).map((_, col) => (
            <div
              key={col}
              className="relative bg-gray-100 border-r border-b"
              style={{ width: `${columnWidths[col]}px` }}
            >
              <div className="flex items-center justify-center h-8">
                {indexToColumn(col)}
              </div>
              <div
                className="absolute top-0 right-0 w-2 h-full cursor-ew-resize hover:bg-blue-200"
                onMouseDown={(e) => handleColumnResize(col, e)}
              />
            </div>
          ))}
        </div>

        {/* Rows */}
        {Array.from({ length: rows }).map((_, rowIndex) => (
          <div key={rowIndex} className="flex">
            {/* Row Header */}
            <div className="relative w-12 bg-gray-100 border-r border-b">
              <div className="flex items-center justify-center h-full">
                {rowIndex + 1}
              </div>
              <div
                className="absolute bottom-0 left-0 h-2 w-full cursor-ns-resize hover:bg-blue-200"
                onMouseDown={(e) => handleRowResize(rowIndex, e)}
              />
            </div>

            {/* Cells */}
            {Array.from({ length: cols }).map((_, colIndex) => {
              const isSelected =
                selectedCell.row === rowIndex && selectedCell.col === colIndex;
              const cellKey = `${rowIndex}-${colIndex}`;
              const isError = errors.has(cellKey);
              const isInRange = selectedRange && 
                rowIndex >= selectedRange.startRow && 
                rowIndex <= selectedRange.endRow && 
                colIndex >= selectedRange.startCol && 
                colIndex <= selectedRange.endCol;

              // Highlight logic for Find & Replace
              const isMatched = findReplace.matches.some(
                (m) => m.row === rowIndex && m.col === colIndex
              );
              const isCurrentMatch =
                findReplace.currentMatch >= 0 &&
                findReplace.currentMatch < findReplace.matches.length &&
                findReplace.matches[findReplace.currentMatch]?.row === rowIndex &&
                findReplace.matches[findReplace.currentMatch]?.col === colIndex;

              return (
                <div
                  key={colIndex}
                  className={classNames(
                    "relative border-r border-b",
                    {
                      "bg-blue-50": isSelected,
                      "bg-blue-100": isInRange && !isSelected,
                      "bg-white": !isSelected && !isInRange,
                      "bg-red-100": isError,
                      "ring-1 ring-yellow-400": isMatched,
                      "ring-2 ring-green-500": isCurrentMatch
                    }
                  )}
                  style={{
                    width: `${columnWidths[colIndex]}px`,
                    height: `${rowHeights[rowIndex]}px`,
                  }}
                  draggable
                  onDragStart={(e) => handleDragStart(e, rowIndex, colIndex)}
                  onDragOver={handleDragOver}
                  onDrop={(e) => handleDrop(e, rowIndex, colIndex)}
                  onClick={(e) => handleCellClick(rowIndex, colIndex, e.shiftKey)}
                >
                  <div className="relative w-full h-full">
                    {/* Display the evaluated value */}
                    <div 
                      className={classNames(
                        "absolute inset-0 px-2 flex items-center",
                        {
                          "font-bold": boldCells.has(cellKey),
                          "italic": italicCells.has(cellKey)
                        }
                      )}
                      style={{
                        fontSize: `${fontSize}px`,
                        color: color,
                        pointerEvents: "none"
                      }}
                    >
                      {cells[rowIndex][colIndex]?.startsWith('=') 
                        ? evaluatedValues[rowIndex][colIndex] 
                        : null}
                    </div>
                    
                    {/* Input for editing */}
                    <input
                      type="text"
                      value={cells[rowIndex]?.[colIndex] || ""}
                      onChange={(e) => {
                        handleCellChange(rowIndex, colIndex, e.target.value);
                      }}
                      className={classNames(
                        "w-full h-full px-2 outline-none bg-transparent",
                        {
                          "font-bold": boldCells.has(cellKey),
                          "italic": italicCells.has(cellKey)
                        }
                      )}
                      style={{
                        fontSize: `${fontSize}px`,
                        color: cells[rowIndex][colIndex]?.startsWith('=') ? 'transparent' : color,
                      }}
                    />
                  </div>
                </div>
              );
            })}
          </div>
        ))}
      </div>

      {/* Modals */}
      {findReplace.show && (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
          <div className="bg-white p-4 rounded w-96">
            <div className="mb-4">
              <input
                type="text"
                placeholder="Find"
                value={findReplace.find}
                onChange={(e) =>
                  setFindReplace((p) => ({ ...p, find: e.target.value }))
                }
                className="w-full p-2 border rounded mb-2"
              />
              <input
                type="text"
                placeholder="Replace with"
                value={findReplace.replace}
                onChange={(e) =>
                  setFindReplace((p) => ({ ...p, replace: e.target.value }))
                }
                className="w-full p-2 border rounded"
              />
            </div>
            <div className="flex justify-between items-center mb-2">
              <span>
                {findReplace.matches.length > 0
                  ? `${findReplace.currentMatch + 1}/${
                      findReplace.matches.length
                    } matches`
                  : "No matches found"}
              </span>
              <div className="flex gap-2">
                <button
                  onClick={findMatches}
                  className="px-2 py-1 bg-blue-500 text-white rounded"
                >
                  Find
                </button>
              </div>
            </div>
            <div className="flex gap-2">
              <button
                onClick={handleReplace}
                className="px-2 py-1 bg-blue-500 text-white rounded flex-1"
                disabled={!findReplace.matches.length}
              >
                Replace
              </button>
              <button
                onClick={handleReplaceAll}
                className="px-2 py-1 bg-purple-500 text-white rounded flex-1"
                disabled={!findReplace.matches.length}
              >
                Replace All
              </button>
              <button
                onClick={() =>
                  setFindReplace((p) => ({ ...p, show: false, matches: [] }))
                }
                className="px-2 py-1 bg-gray-500 text-white rounded"
              >
                Close
              </button>
            </div>
          </div>
        </div>
      )}

      {sortConfig.show && (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
          <div className="bg-white p-4 rounded w-96">
            <h3 className="text-lg font-semibold mb-4">Sort Data</h3>
            <div className="mb-4">
              <label className="block mb-2">Sort by column:</label>
              <select
                value={sortConfig.column}
                onChange={(e) => setSortConfig((p) => ({ ...p, column: parseInt(e.target.value) }))}
                className="w-full p-2 border rounded mb-2"
              >
                {Array.from({ length: cols }).map((_, idx) => (
                  <option key={idx} value={idx}>
                    {indexToColumn(idx)}
                  </option>
                ))}
              </select>
              
              <div className="flex gap-4 mt-2">
                <label className="flex items-center">
                  <input
                    type="radio"
                    name="sortDirection"
                    checked={sortConfig.direction === 'asc'}
                    onChange={() => setSortConfig((p) => ({ ...p, direction: 'asc' }))}
                    className="mr-2"
                  />
                  Ascending
                </label>
                <label className="flex items-center">
                  <input
                    type="radio"
                    name="sortDirection"
                    checked={sortConfig.direction === 'desc'}
                    onChange={() => setSortConfig((p) => ({ ...p, direction: 'desc' }))}
                    className="mr-2"
                  />
                  Descending
                </label>
              </div>
            </div>
            <div className="flex gap-2">
              <button
                onClick={sortData}
                className="px-2 py-1 bg-blue-500 text-white rounded flex-1"
              >
                Sort
              </button>
              <button
                onClick={() => setSortConfig((p) => ({ ...p, show: false }))}
                className="px-2 py-1 bg-gray-500 text-white rounded"
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}

      {chartData && (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
          <div className="bg-white p-4 rounded w-3/4">
            <Bar data={chartData} options={{ responsive: true }} />
            <button
              onClick={() => setChartData(null)}
              className="mt-2 px-2 py-1 bg-gray-500 text-white rounded"
            >
              Close
            </button>
          </div>
        </div>
      )}
    </div>
  );
};

export default Spreadsheet;