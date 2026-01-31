// ignore_for_file: avoid_print

import 'package:a1/a1.dart';
import 'package:worksheet_formula/worksheet_formula.dart';

/// A simple in-memory evaluation context backed by a Map.
class MapEvaluationContext implements EvaluationContext {
  final Map<A1, FormulaValue> cells;
  final FunctionRegistry _registry;
  final A1 _currentCell;

  MapEvaluationContext(
    this._registry, {
    Map<A1, FormulaValue>? cells,
    A1? currentCell,
  })  : cells = cells ?? {},
        _currentCell = currentCell ?? 'A1'.a1;

  @override
  A1 get currentCell => _currentCell;

  @override
  String? get currentSheet => null;

  @override
  bool get isCancelled => false;

  @override
  FormulaValue getCellValue(A1 cell) => cells[cell] ?? const EmptyValue();

  @override
  FormulaValue getRangeValues(A1Range range) {
    final from = range.from.a1;
    final to = range.to.a1;
    if (from == null || to == null) {
      return const FormulaValue.error(FormulaError.ref);
    }
    final rows = <List<FormulaValue>>[];
    for (var row = from.row; row <= to.row; row++) {
      final rowValues = <FormulaValue>[];
      for (var col = from.column; col <= to.column; col++) {
        final cell = A1.fromVector(col, row);
        rowValues.add(cells[cell] ?? const EmptyValue());
      }
      rows.add(rowValues);
    }
    return FormulaValue.range(rows);
  }

  @override
  FormulaFunction? getFunction(String name) => _registry.get(name);
}

// -- Custom function example --------------------------------------------------

/// DOUBLE(number) - Doubles a number.
class DoubleFunction extends FormulaFunction {
  @override
  String get name => 'DOUBLE';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final n = value.toNumber();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    return FormulaValue.number(n * 2);
  }
}

// -- Main ---------------------------------------------------------------------

void main() {
  final engine = FormulaEngine();

  // 1. Basic parsing and evaluation
  print('--- Basic Parsing & Evaluation ---');
  final context1 = MapEvaluationContext(engine.functions);

  final result1 = engine.evaluateString('=1+2*3', context1);
  print('=1+2*3  =>  $result1'); // 7

  final result2 = engine.evaluateString('=(1+2)*3', context1);
  print('=(1+2)*3  =>  $result2'); // 9

  // 2. Cell references
  print('\n--- Cell References ---');
  final context2 = MapEvaluationContext(
    engine.functions,
    cells: {
      'A1'.a1: const NumberValue(10),
      'A2'.a1: const NumberValue(20),
      'A3'.a1: const NumberValue(30),
    },
  );

  print('A1=10, A2=20, A3=30');
  print('=A1+A2+A3  =>  ${engine.evaluateString('=A1+A2+A3', context2)}');
  print('=SUM(A1:A3)  =>  ${engine.evaluateString('=SUM(A1:A3)', context2)}');
  print(
      '=AVERAGE(A1:A3)  =>  ${engine.evaluateString('=AVERAGE(A1:A3)', context2)}');

  // 3. Dependency graph for recalculation
  print('\n--- Dependency Graph ---');
  final graph = DependencyGraph();
  final formulas = <A1, String>{
    'B1'.a1: '=A1+1',
    'C1'.a1: '=B1*2',
  };

  // Register dependencies from formulas
  for (final entry in formulas.entries) {
    final refs = engine.getCellReferences(entry.value);
    graph.updateDependencies(entry.key, refs);
  }

  // Simulate: A1 changes -> which cells need recalculation?
  final toRecalc = graph.getCellsToRecalculate('A1'.a1);
  print('When A1 changes, recalculate: $toRecalc');

  // Evaluate the chain with A1=10
  final cells = <A1, FormulaValue>{'A1'.a1: const NumberValue(10)};
  for (final cell in toRecalc) {
    final formula = formulas[cell]!;
    final ctx = MapEvaluationContext(engine.functions, cells: cells);
    cells[cell] = engine.evaluateString(formula, ctx);
    print('  $cell ($formula)  =>  ${cells[cell]}');
  }

  // 4. Custom function registration
  print('\n--- Custom Functions ---');
  engine.registerFunction(DoubleFunction());
  final context4 = MapEvaluationContext(engine.functions);

  print('=DOUBLE(21)  =>  ${engine.evaluateString('=DOUBLE(21)', context4)}');
  print(
      '=DOUBLE(DOUBLE(5))  =>  ${engine.evaluateString('=DOUBLE(DOUBLE(5))', context4)}');

  // 5. Conditional logic
  print('\n--- Conditional Logic ---');
  final context5 = MapEvaluationContext(
    engine.functions,
    cells: {'A1'.a1: const NumberValue(42)},
  );

  print('A1=42');
  print(
      '=IF(A1>10,"big","small")  =>  ${engine.evaluateString('=IF(A1>10,"big","small")', context5)}');
  print(
      '=IF(A1<10,"low","high")  =>  ${engine.evaluateString('=IF(A1<10,"low","high")', context5)}');
  print(
      '=IFERROR(1/0,"oops")  =>  ${engine.evaluateString('=IFERROR(1/0,"oops")', context5)}');
}
