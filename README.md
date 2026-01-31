# worksheet_formula

A standalone formula engine for spreadsheet-like calculations in Dart.

## Features

- Excel/Google Sheets compatible formula parsing
- 43 built-in functions (math, logical, text, statistical, lookup, date)
- Type-safe formula values with Excel-compatible error handling
- Cell dependency tracking for efficient recalculation
- Custom function registration
- Parse caching for performance
- Zero UI dependencies -- works with any data source

## Installation

```yaml
dependencies:
  worksheet_formula: ^0.2.0
```

## Quick Start

```dart
import 'package:worksheet_formula/worksheet_formula.dart';

// Create the engine
final engine = FormulaEngine();

// Parse and evaluate
final ast = engine.parse('=1+2*3');
final result = engine.evaluate(ast, myContext);
// result = NumberValue(7)
```

To connect your data, implement `EvaluationContext`:

```dart
class MyContext implements EvaluationContext {
  final Map<A1, FormulaValue> cells;
  final FunctionRegistry registry;

  MyContext(this.registry, this.cells);

  @override
  A1 get currentCell => 'A1'.a1;
  @override
  String? get currentSheet => null;
  @override
  bool get isCancelled => false;

  @override
  FormulaValue getCellValue(A1 cell) => cells[cell] ?? const EmptyValue();

  @override
  FormulaValue getRangeValues(A1Range range) {
    // Build 2D matrix from your data source
    // ...
  }

  @override
  FormulaFunction? getFunction(String name) => registry.get(name);
}
```

Then evaluate formulas with cell references:

```dart
final context = MyContext(engine.functions, {
  'A1'.a1: NumberValue(10),
  'A2'.a1: NumberValue(20),
});
final result = engine.evaluateString('=SUM(A1:A2)', context);
// result = NumberValue(30)
```

## Built-in Functions

### Math (10)
`SUM`, `AVERAGE`, `MIN`, `MAX`, `ABS`, `ROUND`, `INT`, `MOD`, `SQRT`, `POWER`

### Logical (8)
`IF`, `AND`, `OR`, `NOT`, `IFERROR`, `IFNA`, `TRUE`, `FALSE`

### Text (10)
`CONCAT`, `CONCATENATE`, `LEFT`, `RIGHT`, `MID`, `LEN`, `LOWER`, `UPPER`, `TRIM`, `TEXT`

### Statistical (6)
`COUNT`, `COUNTA`, `COUNTBLANK`, `COUNTIF`, `SUMIF`, `AVERAGEIF`

### Lookup (3)
`VLOOKUP`, `INDEX`, `MATCH`

### Date (6)
`DATE`, `TODAY`, `NOW`, `YEAR`, `MONTH`, `DAY`

## Custom Functions

```dart
class DiscountFunction extends FormulaFunction {
  @override String get name => 'DISCOUNT';
  @override int get minArgs => 2;
  @override int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final price = values[0].toNumber() ?? 0;
    final rate = values[1].toNumber() ?? 0;
    return FormulaValue.number(price * (1 - rate));
  }
}

engine.registerFunction(DiscountFunction());
engine.evaluateString('=DISCOUNT(100, 0.2)', context);
// result = NumberValue(80)
```

## Dependency Tracking

Track which cells depend on which, and determine recalculation order:

```dart
final graph = DependencyGraph();

// B1 = A1 + 1
graph.updateDependencies('B1'.a1, {'A1'.a1});
// C1 = B1 * 2
graph.updateDependencies('C1'.a1, {'B1'.a1});

// When A1 changes, recalculate in order:
final toRecalc = graph.getCellsToRecalculate('A1'.a1);
// [B1, C1]

// Detect circular references
graph.hasCircularReference('A1'.a1); // false
```

## TEXT Format Codes

The `TEXT` function supports Excel-style format codes:

| Format | Example | Result |
|--------|---------|--------|
| `0.00` | `TEXT(3.14159, "0.00")` | `3.14` |
| `#,##0` | `TEXT(1234567, "#,##0")` | `1,234,567` |
| `0%` | `TEXT(0.75, "0%")` | `75%` |
| `000` | `TEXT(5, "000")` | `005` |
| `0.0E+0` | `TEXT(1234, "0.0E+0")` | `1.2E+3` |

## Supported Operators

| Operator | Description | Example |
|----------|-------------|---------|
| `+` `-` `*` `/` | Arithmetic | `=A1+B1*2` |
| `^` | Power | `=2^10` |
| `&` | Concatenation | `="Hello"&" "&"World"` |
| `=` `<>` `<` `>` `<=` `>=` | Comparison | `=A1>10` |
| `-` (prefix) | Negation | `=-A1` |
| `%` (postfix) | Percent | `=50%` |

## Error Types

All Excel-compatible error types are supported: `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A`, `#NULL!`, `#CALC!`, `#CIRCULAR!`

## License

See [LICENSE](LICENSE) file.
