const { DepParser } = require('./index.js');

// Test the DepParser with the problematic formula
const formula = 'IFERROR(X7 / DetailedCapTableTotalOutstandingShares, 0)';
const position = { row: 7, col: 26, sheet: 'Detailed Cap Table' };

console.log('Testing formula:', formula);
console.log('Position:', position);
console.log('');

// Test 1: Without onVariable callback
console.log('=== Test 1: Without onVariable callback ===');
const parser1 = new DepParser();
try {
    const refs1 = parser1.parse(formula, position);
    console.log('References found:', JSON.stringify(refs1, null, 2));
} catch (e) {
    console.log('Error:', e.message);
}
console.log('');

// Test 2: With onVariable callback
console.log('=== Test 2: With onVariable callback ===');
const variableNames = new Set(['DetailedCapTableTotalOutstandingShares', 'X7']);

const parser2 = new DepParser({
    onVariable: (name, sheet) => {
        console.log(`onVariable called: name="${name}", sheet="${sheet}"`);
        if (variableNames.has(name)) {
            console.log(`  -> Recognized as variable, returning marker`);
            return { name };
        }
        console.log(`  -> Not in variable list, returning null`);
        return null;
    }
});

try {
    const refs2 = parser2.parse(formula, position);
    console.log('References found:', JSON.stringify(refs2, null, 2));
} catch (e) {
    console.log('Error:', e.message);
    console.log('Stack:', e.stack);
}
console.log('');

// Test 3: Simpler formula to understand parsing
console.log('=== Test 3: Simpler formula "X7 + Y5" ===');
const parser3 = new DepParser();
try {
    const refs3 = parser3.parse('X7 + Y5', position);
    console.log('References found:', JSON.stringify(refs3, null, 2));
} catch (e) {
    console.log('Error:', e.message);
}
console.log('');

// Test 4: Test with just the variable
console.log('=== Test 4: Just the variable "DetailedCapTableTotalOutstandingShares" ===');
const parser4 = new DepParser({
    onVariable: (name, sheet) => {
        console.log(`onVariable called: name="${name}"`);
        return { name };
    }
});
try {
    const refs4 = parser4.parse('DetailedCapTableTotalOutstandingShares', position);
    console.log('References found:', JSON.stringify(refs4, null, 2));
} catch (e) {
    console.log('Error:', e.message);
}
