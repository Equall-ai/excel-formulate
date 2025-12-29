const expect = require('chai').expect;
const { evaluateWorkbook } = require('../../grammar/workbook/evaluator');

describe('Workbook Evaluation', function () {
    describe('Variable evaluation with formulas', function () {
        it('should evaluate variables and use them in formulas', function () {
            const workbook = {
                variables: [
                    {
                        name: 'TotalShares',
                        formula: "'Data'!$A$1",
                        value: undefined,
                        result: undefined
                    },
                    {
                        name: 'SharePrice',
                        formula: "'Data'!$B$1",
                        value: undefined,
                        result: undefined
                    }
                ],
                sheets: {
                    'Data': {
                        'A1': {
                            type: 'constant',
                            address: { row: 1, col: 1, sheet: 'Data' },
                            value: 1000
                        },
                        'B1': {
                            type: 'constant',
                            address: { row: 1, col: 2, sheet: 'Data' },
                            value: 50
                        }
                    },
                    'Calculations': {
                        'C1': {
                            type: 'formula',
                            address: { row: 1, col: 3, sheet: 'Calculations' },
                            formulaData: {
                                formulaString: '=TotalShares * SharePrice',
                                cachesResult: undefined
                            },
                            value: undefined
                        }
                    }
                }
            };

            const result = evaluateWorkbook(workbook);

            // Verify variable TotalShares was evaluated
            const totalSharesVar = result.variables.find(v => v.name === 'TotalShares');
            expect(totalSharesVar).to.exist;
            expect(totalSharesVar.result).to.equal(1000);

            // Verify variable SharePrice was evaluated
            const sharePriceVar = result.variables.find(v => v.name === 'SharePrice');
            expect(sharePriceVar).to.exist;
            expect(sharePriceVar.result).to.equal(50);

            // Verify calculation cell was evaluated correctly using variables
            const calcCell = result.sheets['Calculations']['C1'];
            expect(calcCell).to.exist;
            expect(calcCell.value).to.equal(50000);
        });

        it('should handle variables with simple numeric values', function () {
            const workbook = {
                variables: [
                    {
                        name: 'Rate',
                        formula: "'Settings'!$A$1",
                        value: undefined,
                        result: undefined
                    }
                ],
                sheets: {
                    'Settings': {
                        'A1': {
                            type: 'constant',
                            address: { row: 1, col: 1, sheet: 'Settings' },
                            value: 0.05
                        }
                    },
                    'Calculations': {
                        'B1': {
                            type: 'formula',
                            address: { row: 1, col: 2, sheet: 'Calculations' },
                            formulaData: {
                                formulaString: '=100 * Rate',
                                cachesResult: undefined
                            },
                            value: undefined
                        }
                    }
                }
            };

            const result = evaluateWorkbook(workbook);

            const rateVar = result.variables.find(v => v.name === 'Rate');
            expect(rateVar.result).to.equal(0.05);

            const calcCell = result.sheets['Calculations']['B1'];
            expect(calcCell.value).to.equal(5);
        });

        it('should handle multiple variables with dependencies', function () {
            const workbook = {
                variables: [
                    {
                        name: 'BaseValue',
                        formula: "'Data'!$A$1",
                        value: undefined,
                        result: undefined
                    },
                    {
                        name: 'Multiplier',
                        formula: "'Data'!$B$1",
                        value: undefined,
                        result: undefined
                    }
                ],
                sheets: {
                    'Data': {
                        'A1': {
                            type: 'constant',
                            address: { row: 1, col: 1, sheet: 'Data' },
                            value: 10
                        },
                        'B1': {
                            type: 'constant',
                            address: { row: 1, col: 2, sheet: 'Data' },
                            value: 3
                        }
                    },
                    'Results': {
                        'C1': {
                            type: 'formula',
                            address: { row: 1, col: 3, sheet: 'Results' },
                            formulaData: {
                                formulaString: '=(BaseValue + Multiplier) * 2',
                                cachesResult: undefined
                            },
                            value: undefined
                        }
                    }
                }
            };

            const result = evaluateWorkbook(workbook);

            const resultCell = result.sheets['Results']['C1'];
            expect(resultCell.value).to.equal(26);
        });

        it('should handle cells with formulas referencing each other', function () {
            const workbook = {
                variables: [],
                sheets: {
                    'Sheet1': {
                        'A1': {
                            type: 'constant',
                            address: { row: 1, col: 1, sheet: 'Sheet1' },
                            value: 100
                        },
                        'A2': {
                            type: 'formula',
                            address: { row: 2, col: 1, sheet: 'Sheet1' },
                            formulaData: {
                                formulaString: '=A1 * 2',
                                cachesResult: undefined
                            },
                            value: undefined
                        },
                        'A3': {
                            type: 'formula',
                            address: { row: 3, col: 1, sheet: 'Sheet1' },
                            formulaData: {
                                formulaString: '=A2 + 50',
                                cachesResult: undefined
                            },
                            value: undefined
                        }
                    }
                }
            };

            const result = evaluateWorkbook(workbook);

            expect(result.sheets['Sheet1']['A2'].value).to.equal(200);
            expect(result.sheets['Sheet1']['A3'].value).to.equal(250);
        });
    });
});
