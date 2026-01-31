# Worksheet Formula Integration Example

A Flutter app demonstrating how to integrate the [`worksheet_formula`](https://pub.dev/packages/worksheet_formula) engine with the [`worksheet`](https://pub.dev/packages/worksheet) widget.

## What This Shows

- **Formula evaluation in a live spreadsheet** -- cells containing formulas like `=B2*C2` and `=SUM(D2:D4)` display computed results
- **Dependency tracking** -- when a cell changes, only affected formulas are recalculated
- **Custom functions** -- a `DISCOUNT(price, rate)` function registered with the engine
- **Circular reference detection** -- self-referencing formulas return `#CIRCULAR!`

## Architecture

```
┌──────────────────────┐     ┌─────────────────────┐
│   SparseWorksheetData│     │    FormulaEngine     │
│   (raw cell storage) │     │  (parser + registry) │
└──────────┬───────────┘     └──────────┬──────────┘
           │                            │
           ▼                            ▼
┌──────────────────────────────────────────────────┐
│              FormulaWorksheetData                 │
│                                                   │
│  Wraps WorksheetData, intercepts getCell() to     │
│  evaluate formula cells. Caches results and uses  │
│  DependencyGraph to invalidate on changes.        │
└──────────────────────┬───────────────────────────┘
                       │
                       ▼
┌──────────────────────────────────────────────────┐
│              Worksheet widget                     │
│  Renders the grid, displays computed values       │
└──────────────────────────────────────────────────┘
```

### Key Classes

**`FormulaWorksheetData`** (`lib/src/formula_worksheet_data.dart`)

Implements `WorksheetData` by wrapping an inner data source. When `getCell()` encounters a formula cell, it evaluates the formula and returns the computed `CellValue`. Results are cached and invalidated via `DependencyGraph` when upstream cells change.

**`WorksheetEvaluationContext`** (`lib/src/worksheet_evaluation_context.dart`)

Implements `EvaluationContext` to bridge `WorksheetData` cell values into the formula engine's `FormulaValue` type system. Handles recursive evaluation of formula cells that reference other formula cells, with circular reference detection.

## Running

```bash
cd example/worksheet_integration
flutter pub get
flutter run
```

## Demo Spreadsheet

The example creates a small invoice:

|   | A       | B     | C   | D                    |
|---|---------|-------|-----|----------------------|
| 1 | Item    | Price | Qty | Total                |
| 2 | Apples  | 1.50  | 10  | `=B2*C2` → 15       |
| 3 | Oranges | 2.00  | 5   | `=B3*C3` → 10       |
| 4 | Bananas | 0.75  | 12  | `=B4*C4` → 9        |
| 6 |         |       | Subtotal: | `=SUM(D2:D4)` → 34 |
| 7 |         |       | Tax (10%): | `=D6*0.1` → 3.4   |
| 8 |         |       | Total: | `=D6+D7` → 37.4     |
| 10 |        |       | With 15% discount: | `=DISCOUNT(D8, 0.15)` → 31.79 |

Tap any cell to see its formula and computed result in the info bar.

## Adapting for Your App

1. Create your `SparseWorksheetData` with formula cells using `'=SUM(A1:A10)'.formula`
2. Wrap it with `FormulaWorksheetData(data, engine: engine)`
3. Pass the wrapper to the `Worksheet` widget
4. Register custom functions on the `FormulaEngine` before wrapping

```dart
final engine = FormulaEngine();
engine.registerFunction(MyCustomFunction());

final rawData = SparseWorksheetData(
  rowCount: 100,
  columnCount: 26,
  cells: {
    (0, 0): Cell.number(10),
    (0, 1): '=A1*2'.formula,
  },
);

final data = FormulaWorksheetData(rawData, engine: engine);

Worksheet(data: data, rowCount: 100, columnCount: 26)
```
