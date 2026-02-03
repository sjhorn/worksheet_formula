// ignore_for_file: avoid_print

import 'package:a1/a1.dart';
import 'package:worksheet_formula/src/evaluation/context.dart';
import 'package:worksheet_formula/src/evaluation/errors.dart';
import 'package:worksheet_formula/src/evaluation/value.dart';
import 'package:worksheet_formula/src/formula_engine.dart';
import 'package:worksheet_formula/src/functions/function.dart';
import 'package:worksheet_formula/src/functions/registry.dart';

class _BenchContext implements EvaluationContext {
  final Map<A1, FormulaValue> cells;
  final FunctionRegistry _registry;

  _BenchContext(this._registry, this.cells);

  @override
  A1 get currentCell => 'A1'.a1;
  @override
  String? get currentSheet => null;
  @override
  FormulaValue? getVariable(String name) => null;
  @override
  bool get isCancelled => false;

  @override
  FormulaValue getCellValue(A1 cell) =>
      cells[cell] ?? const EmptyValue();

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

void main() {
  final engine = FormulaEngine();

  print('=== Evaluate Benchmarks ===\n');

  // Simple arithmetic
  final simpleAst = engine.parse('=1+2*3');
  final emptyContext = _BenchContext(engine.functions, {});
  _benchmark('Simple arithmetic (1+2*3)',
      () => engine.evaluate(simpleAst, emptyContext));

  // Cell references
  final cells10 = <A1, FormulaValue>{
    for (var i = 1; i <= 10; i++)
      A1.fromVector(0, i - 1): FormulaValue.number(i),
  };
  final refAst = engine.parse('=A1+A2+A3+A4+A5');
  final refContext = _BenchContext(engine.functions, cells10);
  _benchmark('Cell references (A1+...+A5)',
      () => engine.evaluate(refAst, refContext));

  // SUM over range
  final sumAst = engine.parse('=SUM(A1:A10)');
  _benchmark('SUM(A1:A10) - 10 cells',
      () => engine.evaluate(sumAst, refContext));

  // SUM over larger range
  final cells100 = <A1, FormulaValue>{
    for (var i = 0; i < 100; i++)
      A1.fromVector(0, i): FormulaValue.number(i),
  };
  final sumLargeAst = engine.parse('=SUM(A1:A100)');
  final largeContext = _BenchContext(engine.functions, cells100);
  _benchmark('SUM(A1:A100) - 100 cells',
      () => engine.evaluate(sumLargeAst, largeContext));

  // IF with comparison
  final ifAst = engine.parse('=IF(A1>5,"big","small")');
  _benchmark('IF(A1>5,"big","small")',
      () => engine.evaluate(ifAst, refContext));

  // Nested functions
  final nestedAst = engine.parse('=ROUND(AVERAGE(A1:A10),2)');
  _benchmark('ROUND(AVERAGE(A1:A10),2)',
      () => engine.evaluate(nestedAst, refContext));
}

void _benchmark(String label, void Function() fn,
    [int iterations = 10000]) {
  for (var i = 0; i < 100; i++) {
    fn();
  }

  final sw = Stopwatch()..start();
  for (var i = 0; i < iterations; i++) {
    fn();
  }
  sw.stop();

  final opsPerSec = (iterations / sw.elapsedMicroseconds * 1000000).round();
  print('  $label');
  print('    ${sw.elapsedMilliseconds}ms for $iterations ops'
      ' ($opsPerSec ops/sec)\n');
}
