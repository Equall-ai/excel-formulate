export interface VariableData {
  type: 'variable';
  name: string;
  formula?: string;
  value?: any;
  result?: any;
}

export interface CellAddress {
  row: number;
  col: number;
  sheet: string;
}

export interface FormulaData {
  formulaString: string;
  cachesResult?: any;
}

export interface CellData {
  type: 'formula' | 'constant';
  address: CellAddress;
  formulaData?: FormulaData;
  value?: any;
}

export interface WorkbookData {
  variables: VariableData[];
  sheets: {
    [sheetName: string]: {
      [address: string]: CellData;
    };
  };
}

export interface DependencyNode {
  id: string;
  dependencyIds: Set<string>;
}

export interface TopologicallySortedNodes {
  topologicallySortedNodes: DependencyNode[];
  recursiveNodes: DependencyNode[];
}

export function evaluateFormula(
  formulaParser: any,
  formula: string,
  context?: { row: number; col: number; sheet: string }
): any;

export function evaluateWorkbook(workbook: WorkbookData): WorkbookData;
