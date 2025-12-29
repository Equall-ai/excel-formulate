export * from './grammar/workbook/evaluator';

export class FormulaParser {
  constructor(options?: {
    onCell?: (ref: { row: number; col: number; sheet: string }) => any;
    onRange?: (ref: { from: { row: number; col: number }; to: { row: number; col: number }; sheet: string }) => any;
    onVariable?: (name: string) => any;
    functions?: Record<string, Function>;
    functionsNeedContext?: Record<string, boolean>;
  });
  parse(formula: string, position?: { row: number; col: number; sheet?: string }): any;
  supportedFunctions(): string[];
}

export class DepParser {
  constructor(config?: {
    onVariable?: (name: string, sheet?: string) => any;
  });
  parse(formula: string, position?: { row: number; col: number; sheet?: string }): any;
  data: any[];
}

export class FormulaError extends Error {
  constructor(message: string);
}

export const SSF: any;

export interface FormulaHelpers {
  argsConvert(args: any[]): any[];
  numToDate(num: number): Date;
  dateToNum(date: Date): number;
}

export const evaluateWorkbook: any;
export const evaluateFormula: any;

export default FormulaParser;
