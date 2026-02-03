# worksheet_formula

A standalone formula engine for spreadsheet-like calculations in Dart.

## Features

- Excel/Google Sheets compatible formula parsing
- 239 built-in functions across 9 categories (math, logical, text, statistical, lookup, date, information, array, financial)
- Dynamic array functions (FILTER, SORT, UNIQUE, SEQUENCE, etc.)
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

### Math & Trigonometry (47)
`SUM`, `AVERAGE`, `MIN`, `MAX`, `ABS`, `ROUND`, `INT`, `MOD`, `SQRT`, `POWER`, `SUMPRODUCT`, `ROUNDUP`, `ROUNDDOWN`, `CEILING`, `FLOOR`, `SIGN`, `PRODUCT`, `RAND`, `RANDBETWEEN`, `PI`, `LN`, `LOG`, `LOG10`, `EXP`, `SIN`, `COS`, `TAN`, `ASIN`, `ACOS`, `ATAN`, `ATAN2`, `DEGREES`, `RADIANS`, `EVEN`, `ODD`, `GCD`, `LCM`, `TRUNC`, `MROUND`, `QUOTIENT`, `COMBIN`, `COMBINA`, `FACT`, `FACTDOUBLE`, `SUMSQ`, `SUBTOTAL`, `AGGREGATE`

### Logical (11)
`IF`, `AND`, `OR`, `NOT`, `IFERROR`, `IFNA`, `TRUE`, `FALSE`, `IFS`, `SWITCH`, `XOR`

### Text (31)
`CONCAT`, `CONCATENATE`, `LEFT`, `RIGHT`, `MID`, `LEN`, `LOWER`, `UPPER`, `TRIM`, `TEXT`, `FIND`, `SEARCH`, `SUBSTITUTE`, `REPLACE`, `VALUE`, `TEXTJOIN`, `PROPER`, `EXACT`, `REPT`, `CHAR`, `CODE`, `CLEAN`, `DOLLAR`, `FIXED`, `T`, `NUMBERVALUE`, `UNICHAR`, `UNICODE`, `TEXTBEFORE`, `TEXTAFTER`, `TEXTSPLIT`

### Statistical (35)
`COUNT`, `COUNTA`, `COUNTBLANK`, `COUNTIF`, `SUMIF`, `AVERAGEIF`, `SUMIFS`, `COUNTIFS`, `AVERAGEIFS`, `MEDIAN`, `MODE.SNGL`, `MODE`, `LARGE`, `SMALL`, `RANK.EQ`, `RANK`, `STDEV.S`, `STDEV.P`, `VAR.S`, `VAR.P`, `PERCENTILE.INC`, `PERCENTILE.EXC`, `PERCENTRANK.INC`, `PERCENTRANK.EXC`, `RANK.AVG`, `FREQUENCY`, `AVEDEV`, `AVERAGEA`, `MAXA`, `MINA`, `TRIMMEAN`, `GEOMEAN`, `HARMEAN`, `MAXIFS`, `MINIFS`

### Lookup & Reference (18)
`VLOOKUP`, `INDEX`, `MATCH`, `HLOOKUP`, `LOOKUP`, `CHOOSE`, `XMATCH`, `XLOOKUP`, `ROW`, `COLUMN`, `ROWS`, `COLUMNS`, `ADDRESS`, `INDIRECT`, `OFFSET`, `TRANSPOSE`, `HYPERLINK`, `AREAS`

### Date/Time (25)
`DATE`, `TODAY`, `NOW`, `YEAR`, `MONTH`, `DAY`, `DAYS`, `DATEDIF`, `DATEVALUE`, `WEEKDAY`, `HOUR`, `MINUTE`, `SECOND`, `TIME`, `EDATE`, `EOMONTH`, `TIMEVALUE`, `WEEKNUM`, `ISOWEEKNUM`, `NETWORKDAYS`, `NETWORKDAYS.INTL`, `WORKDAY`, `WORKDAY.INTL`, `DAYS360`, `YEARFRAC`

### Information (15)
`ISBLANK`, `ISERROR`, `ISNUMBER`, `ISTEXT`, `ISLOGICAL`, `ISNA`, `TYPE`, `ISERR`, `ISNONTEXT`, `ISEVEN`, `ISODD`, `ISREF`, `N`, `NA`, `ERROR.TYPE`

### Dynamic Array (17)
`SEQUENCE`, `RANDARRAY`, `TOCOL`, `TOROW`, `WRAPROWS`, `WRAPCOLS`, `CHOOSEROWS`, `CHOOSECOLS`, `DROP`, `TAKE`, `EXPAND`, `HSTACK`, `VSTACK`, `FILTER`, `UNIQUE`, `SORT`, `SORTBY`

### Financial (40)
`PMT`, `FV`, `PV`, `NPER`, `RATE`, `IPMT`, `PPMT`, `CUMIPMT`, `CUMPRINC`, `NPV`, `XNPV`, `IRR`, `XIRR`, `MIRR`, `FVSCHEDULE`, `SLN`, `SYD`, `DB`, `DDB`, `VDB`, `PRICE`, `YIELD`, `DURATION`, `MDURATION`, `ACCRINT`, `DISC`, `INTRATE`, `RECEIVED`, `PRICEDISC`, `PRICEMAT`, `TBILLEQ`, `TBILLPRICE`, `TBILLYIELD`, `DOLLARDE`, `DOLLARFR`, `EFFECT`, `NOMINAL`, `PDURATION`, `RRI`, `ISPMT`

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

## Examples

### Standalone Dart

[`example/worksheet_formula_example.dart`](example/worksheet_formula_example.dart) -- a pure Dart example demonstrating parsing, evaluation, cell references, dependency tracking, custom functions, and conditional logic. No Flutter required.

```bash
dart run example/worksheet_formula_example.dart
```

### Flutter + Worksheet Widget

[`example/worksheet_integration/`](example/worksheet_integration/) -- a Flutter app integrating `worksheet_formula` with the [`worksheet`](https://pub.dev/packages/worksheet) widget. Shows formula cells evaluated live in a spreadsheet grid with dependency tracking, caching, and a custom `DISCOUNT` function.

```bash
cd example/worksheet_integration
flutter run
```

See the [integration README](example/worksheet_integration/README.md) for architecture details.

## License

See [LICENSE](LICENSE) file.
