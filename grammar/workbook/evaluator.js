const { FormulaParser } = require('../hooks');
const { DepParser } = require('../dependency/hooks');

const MAX_ITER = 30;

/**
 * @typedef {Object} VariableData
 * @property {'variable'} type
 * @property {string} name
 * @property {string} [formula]
 * @property {any} [value]
 * @property {any} [result]
 */

/**
 * @typedef {Object} CellAddress
 * @property {number} row
 * @property {number} col
 * @property {string} sheet
 */

/**
 * @typedef {Object} FormulaData
 * @property {string} formulaString
 * @property {any} [cachesResult]
 */

/**
 * @typedef {Object} CellData
 * @property {'formula' | 'constant'} type
 * @property {CellAddress} address
 * @property {FormulaData} [formulaData]
 * @property {any} [value]
 */

/**
 * @typedef {Object} WorkbookData
 * @property {VariableData[]} variables
 * @property {{[sheetName: string]: {[address: string]: CellData}}} sheets
 */

/**
 * @typedef {Object} DependencyNode
 * @property {string} id
 * @property {Set<string>} dependencyIds
 */

/**
 * @typedef {Object} TopologicallySortedNodes
 * @property {DependencyNode[]} topologicallySortedNodes
 * @property {DependencyNode[]} recursiveNodes
 */

/**
 * @typedef {Object} DepParserCellReference
 * @property {string} sheet
 * @property {number} row
 * @property {number} col
 */

/**
 * @typedef {Object} DepParserRangeReference
 * @property {string} sheet
 * @property {DepParserCellReference} from
 * @property {DepParserCellReference} to
 */

/**
 * @typedef {Object} DepParserVariableReference
 * @property {string} name
 */

function _rowColToA1(row, col) {
  let colStr = '';
  let c = col;
  while (c > 0) {
    c--;
    colStr = String.fromCharCode((c % 26) + 65) + colStr;
    c = Math.floor(c / 26);
  }
  return `${colStr}${row}`;
}

function _hashCellAddress(address) {
  return `${address.sheet}-${address.row}-${address.col}`;
}

/**
 * @param {DepParserCellReference | DepParserRangeReference | DepParserVariableReference} reference
 * @returns {'cell' | 'range' | 'variable'}
 */
function _getDepParserType(reference) {
  if ('row' in reference && 'col' in reference && 'sheet' in reference) return 'cell';
  if ('from' in reference && 'to' in reference && 'sheet' in reference) return 'range';
  if ('name' in reference) return 'variable';
  return 'cell';
}

/**
 * @param {DepParserRangeReference} reference
 * @param {Map<string, Map<number, Set<number>>>} [cellIndex]
 * @returns {CellAddress[]}
 */
function _convertRangeToCellAddresses(reference, cellIndex) {
  const result = [];
  const MAX_RANGE_SIZE = 1000;

  const rowCount = reference.to.row - reference.from.row + 1;
  const colCount = reference.to.col - reference.from.col + 1;
  const totalCells = rowCount * colCount;

  // If range is large and we have cell index, only include existing cells
  if (totalCells > MAX_RANGE_SIZE && cellIndex) {
    const sheetIndex = cellIndex.get(reference.sheet);
    if (!sheetIndex) return [];

    for (let col = reference.from.col; col <= reference.to.col; col++) {
      const rowsInCol = sheetIndex.get(col);
      if (!rowsInCol) continue;

      for (const row of rowsInCol) {
        if (row >= reference.from.row && row <= reference.to.row) {
          result.push({ row, col, sheet: reference.sheet });
        }
      }
    }
    return result;
  }

  // If range is too large without index, skip silently
  if (totalCells > MAX_RANGE_SIZE) {
    return [];
  }

  // Normal case: small range, include all cells
  for (let row = reference.from.row; row <= reference.to.row; row++) {
    for (let col = reference.from.col; col <= reference.to.col; col++) {
      result.push({ row, col, sheet: reference.sheet });
    }
  }
  return result;
}

/**
 * @param {Array} references
 * @param {string} [defaultSheet]
 * @param {Map<string, Map<number, Set<number>>>} [cellIndex]
 * @returns {string[]}
 */
function _convertDepParserReferenceToId(references, defaultSheet, cellIndex) {
  if (references.length === 0) return [];
  const result = [];
  for (const reference of references) {
    const type = _getDepParserType(reference);
    if (type === 'cell') {
      const cellReference = reference;
      const cellRef = { ...cellReference };
      if (!cellRef.sheet && defaultSheet) cellRef.sheet = defaultSheet;
      result.push(_hashCellAddress(cellRef));
    }
    if (type === 'range') {
      const rangeReference = reference;
      const rangeRef = { ...rangeReference };
      if (!rangeRef.sheet && defaultSheet) rangeRef.sheet = defaultSheet;
      const cellAddresses = _convertRangeToCellAddresses(rangeRef, cellIndex);
      result.push(...cellAddresses.map(cellAddress => _hashCellAddress(cellAddress)));
    }
    if (type === 'variable') {
      const variableReference = reference;
      result.push(variableReference.name);
    }
  }
  return result;
}

/**
 * @param {any} parser - DepParser instance
 * @param {string} id
 * @param {string} formula
 * @param {number} [row]
 * @param {number} [col]
 * @param {string} [sheet]
 * @param {Map<string, Map<number, Set<number>>>} [cellIndex]
 * @returns {DependencyNode}
 */
function _createDependencyNodeWithFormula(parser, id, formula, row, col, sheet, cellIndex) {
  try {
    let formulaToParse = formula;
    if (formulaToParse.startsWith('=')) formulaToParse = formulaToParse.substring(1);

    parser.parse(formulaToParse, { row: row || 1, col: col || 1 });
    const references = parser.data;
    const dependencyIds = _convertDepParserReferenceToId(references, sheet, cellIndex);
    return { id, dependencyIds: new Set(dependencyIds) };
  } catch (error) {
    return { id, dependencyIds: new Set() };
  }
}

/**
 * @param {any} parser - DepParser instance
 * @param {CellData} cell
 * @param {Map<string, Map<number, Set<number>>>} [cellIndex]
 * @returns {DependencyNode}
 */
function _convertCellToDependencyNode(parser, cell, cellIndex) {
  if (cell.type === 'constant' || !cell.formulaData || !cell.formulaData.formulaString) {
    return { id: _hashCellAddress(cell.address), dependencyIds: new Set() };
  }
  return _createDependencyNodeWithFormula(
    parser,
    _hashCellAddress(cell.address),
    cell.formulaData.formulaString,
    cell.address.row,
    cell.address.col,
    cell.address.sheet,
    cellIndex
  );
}

/**
 * @param {any} depParser - DepParser instance
 * @param {VariableData} variable
 * @param {Map<string, Map<number, Set<number>>>} [cellIndex]
 * @returns {DependencyNode}
 */
function _convertVariableToDependencyNode(depParser, variable, cellIndex) {
  if (!variable.formula) return { id: variable.name, dependencyIds: new Set() };
  return _createDependencyNodeWithFormula(depParser, variable.name, variable.formula, 1, 1, undefined, cellIndex);
}

/**
 * @param {DependencyNode[]} dependencies
 * @returns {{[id: string]: string[]}}
 */
function _buildReverseDependencyMappings(dependencies) {
  const reverseDependencyMappings = {};
  for (const dependency of dependencies) {
    for (const id of dependency.dependencyIds) {
      reverseDependencyMappings[id] = [...(reverseDependencyMappings[id] || []), dependency.id];
    }
  }
  return reverseDependencyMappings;
}

/**
 * @param {DependencyNode[]} dependencies
 * @returns {TopologicallySortedNodes}
 */
function _topologicalSorting(dependencies) {
  const indexedDependencies = {};
  const unmodifiedIndexedDependencies = {};
  for (const dependency of dependencies) {
    indexedDependencies[dependency.id] = {
      id: dependency.id,
      dependencyIds: new Set(dependency.dependencyIds)
    };
    unmodifiedIndexedDependencies[dependency.id] = dependency;
  }

  const reverseDependencyMappings = _buildReverseDependencyMappings(dependencies);

  const topologicallySortedNodes = [];
  const recentlyResolvedDependencies = dependencies.filter(dependency => dependency.dependencyIds.size === 0);
  while (recentlyResolvedDependencies.length > 0) {
    const currentDependency = recentlyResolvedDependencies.shift();
    if (!currentDependency) break;
    const currentNode = unmodifiedIndexedDependencies[currentDependency.id];
    topologicallySortedNodes.push(currentNode);
    const dependentIds = reverseDependencyMappings[currentDependency.id];
    if (!dependentIds) continue;

    for (const id of dependentIds) {
      const dependentNode = indexedDependencies[id];
      if (!dependentNode) continue;
      dependentNode.dependencyIds.delete(currentDependency.id);
      if (dependentNode.dependencyIds.size === 0) {
        recentlyResolvedDependencies.push(dependentNode);
      }
    }
  }

  const unresolvedDependencyIds = Object.keys(unmodifiedIndexedDependencies).filter(
    id => indexedDependencies[id].dependencyIds.size > 0
  );
  const recursiveNodes = unresolvedDependencyIds.map(id => unmodifiedIndexedDependencies[id]);
  return { topologicallySortedNodes, recursiveNodes };
}

/**
 * @param {CellData[]} cells
 * @param {VariableData[]} variables
 * @returns {TopologicallySortedNodes}
 */
function _topologicallySortCells(cells, variables) {
  const indexedCells = {};
  const dependencies = [];

  // Create variable index for onVariable callback
  const variableIndex = {};
  for (const variable of variables) {
    variableIndex[variable.name] = variable;
  }

  // DepParser with onVariable callback to recognize variables
  const parser = new DepParser({
    onVariable: (name) => {
      // Return a marker so DepParser recognizes it as a variable reference
      if (variableIndex[name]) {
        return { name };
      }
      return null;
    }
  });

  const cellIndex = new Map();
  for (const cell of cells) {
    const key = _hashCellAddress(cell.address);
    indexedCells[key] = cell;

    if (!cellIndex.has(cell.address.sheet)) {
      cellIndex.set(cell.address.sheet, new Map());
    }
    const sheetIndex = cellIndex.get(cell.address.sheet);
    if (!sheetIndex.has(cell.address.col)) {
      sheetIndex.set(cell.address.col, new Set());
    }
    sheetIndex.get(cell.address.col).add(cell.address.row);
  }

  for (const cell of cells) {
    const cellDependencyNode = _convertCellToDependencyNode(parser, cell, cellIndex);
    dependencies.push(cellDependencyNode);
  }

  for (const variable of variables) {
    const variableDependencyNode = _convertVariableToDependencyNode(parser, variable, cellIndex);
    dependencies.push(variableDependencyNode);
  }

  return _topologicalSorting(dependencies);
}

/**
 * @param {any} formulaParser - FormulaParser instance
 * @param {string} formula
 * @param {{row: number; col: number; sheet: string}} [context]
 * @returns {any}
 */
function evaluateFormula(formulaParser, formula, context) {
  try {
    let formulaStr = formula;
    if (formulaStr.startsWith('=')) formulaStr = formulaStr.substring(1);

    const position = context || { row: 1, col: 1, sheet: 'Sheet1' };
    return formulaParser.parse(formulaStr, position);
  } catch (error) {
    return 0;
  }
}

/**
 * @param {WorkbookData} workbook
 * @returns {WorkbookData}
 */
function evaluateWorkbook(workbook) {
  const allCells = Object.values(workbook.sheets).flatMap(sheet => Object.values(sheet));
  const { topologicallySortedNodes, recursiveNodes } = _topologicallySortCells(allCells, workbook.variables);

  const variableIndex = {};
  const cellIndex = {};

  for (const variable of workbook.variables) {
    variableIndex[variable.name] = variable;
  }

  for (const cell of allCells) {
    cellIndex[_hashCellAddress(cell.address)] = cell;
  }

  const formulaParser = new FormulaParser({
    onCell: (ref) => {
      const address = _rowColToA1(ref.row, ref.col);
      const cell = workbook.sheets[ref.sheet]?.[address];
      if (!cell) {
        return 0;
      }

      if (cell.type === 'formula' && cell.formulaData?.cachesResult !== undefined) {
        return cell.formulaData.cachesResult;
      }

      return cell.value ?? 0;
    },
    onRange: (ref) => {
      const result = [];
      const sheetData = workbook.sheets[ref.sheet];

      if (!sheetData) {
        for (let row = ref.from.row; row <= ref.to.row; row++) {
          result.push(new Array(ref.to.col - ref.from.col + 1).fill(0));
        }
        return result;
      }

      for (let row = ref.from.row; row <= ref.to.row; row++) {
        const rowData = [];
        for (let col = ref.from.col; col <= ref.to.col; col++) {
          const address = _rowColToA1(row, col);
          const cell = sheetData[address];

          if (!cell) {
            rowData.push(0);
          } else if (cell.type === 'formula' && cell.formulaData?.cachesResult !== undefined) {
            rowData.push(cell.formulaData.cachesResult);
          } else {
            rowData.push(cell.value ?? 0);
          }
        }
        result.push(rowData);
      }
      return result;
    },
    onVariable: (name) => {
      const variable = variableIndex[name];
      if (!variable) return 0;

      // Return pre-computed result or value (as plain values, not wrapped)
      return variable.result ?? variable.value ?? 0;
    }
  });

  for (const node of topologicallySortedNodes) {
    const cell = cellIndex[node.id];

    // Check if it's a variable (not in cellIndex)
    if (!cell) {
      const variable = variableIndex[node.id];
      if (variable && variable.formula) {
        const result = evaluateFormula(formulaParser, variable.formula);
        variable.result = result;
      }
      continue;
    }

    // Now handle cells
    if (!cell.formulaData) {
      continue;
    }

    if (cell.formulaData.cachesResult !== undefined && cell.formulaData.cachesResult !== null) {
      cell.value = cell.formulaData.cachesResult;
      continue;
    }

    if (cell.type === 'formula' && cell.formulaData?.formulaString) {
      const result = evaluateFormula(
        formulaParser,
        cell.formulaData.formulaString,
        { row: cell.address.row, col: cell.address.col, sheet: cell.address.sheet }
      );
      cell.value = result;
      cell.formulaData.cachesResult = result;
    }
  }

  let iter = 0;
  while (iter < MAX_ITER) {
    for (const node of recursiveNodes) {
      const cell = cellIndex[node.id];
      if (cell && cell.type === 'formula' && cell.formulaData?.formulaString) {
        const result = evaluateFormula(
          formulaParser,
          cell.formulaData.formulaString,
          { row: cell.address.row, col: cell.address.col, sheet: cell.address.sheet }
        );
        cell.value = result;
        cell.formulaData.cachesResult = result;
      }
    }
    iter++;
  }

  return workbook;
}

module.exports = {
  evaluateWorkbook,
  evaluateFormula
};
