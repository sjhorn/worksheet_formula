import 'package:a1/a1.dart';
import 'package:test/test.dart';
import 'package:worksheet_formula/worksheet_formula.dart';

/// Integration test using the public API via barrel export.
class _SimpleContext implements EvaluationContext {
  final Map<A1, FormulaValue> cells;
  final FunctionRegistry _registry;

  _SimpleContext(this._registry, [Map<A1, FormulaValue>? cells])
      : cells = cells ?? {};

  @override
  A1 get currentCell => 'A1'.a1;
  @override
  String? get currentSheet => null;
  @override
  FormulaValue? getVariable(String name) => null;
  @override
  bool get isCancelled => false;

  @override
  FormulaValue getCellValue(A1 cell) => cells[cell] ?? const EmptyValue();

  @override
  FormulaValue getRangeValues(A1Range range) {
    final from = range.from;
    final to = range.to;
    final fromA1 = from.a1;
    final toA1 = to.a1;
    if (fromA1 == null || toA1 == null) {
      return const FormulaValue.error(FormulaError.ref);
    }
    final rows = <List<FormulaValue>>[];
    for (var row = fromA1.row; row <= toA1.row; row++) {
      final rowValues = <FormulaValue>[];
      for (var col = fromA1.column; col <= toA1.column; col++) {
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
  group('Integration: public API', () {
    test('parse, evaluate, and get references', () {
      final engine = FormulaEngine();
      final context = _SimpleContext(
        engine.functions,
        {
          'A1'.a1: const NumberValue(10),
          'A2'.a1: const NumberValue(20),
          'A3'.a1: const NumberValue(30),
        },
      );

      // Parse and evaluate a simple formula
      final result = engine.evaluateString('=A1+A2+A3', context);
      expect(result, const NumberValue(60));

      // Get cell references
      final refs = engine.getCellReferences('=A1+A2+A3');
      expect(refs, {'A1'.a1, 'A2'.a1, 'A3'.a1});
    });

    test('SUM over a range', () {
      final engine = FormulaEngine();
      final context = _SimpleContext(
        engine.functions,
        {
          'A1'.a1: const NumberValue(1),
          'A2'.a1: const NumberValue(2),
          'A3'.a1: const NumberValue(3),
        },
      );

      final result = engine.evaluateString('=SUM(A1:A3)', context);
      expect(result, const NumberValue(6));
    });

    test('IF with comparison', () {
      final engine = FormulaEngine();
      final context = _SimpleContext(
        engine.functions,
        {'A1'.a1: const NumberValue(42)},
      );

      final result =
          engine.evaluateString('=IF(A1>10,"big","small")', context);
      expect(result, const TextValue('big'));
    });

    test('text functions', () {
      final engine = FormulaEngine();
      final context = _SimpleContext(engine.functions);

      expect(
        engine.evaluateString('=UPPER("hello")', context),
        const TextValue('HELLO'),
      );
      expect(
        engine.evaluateString('=LEN("hello")', context),
        const NumberValue(5),
      );
      expect(
        engine.evaluateString('=CONCAT("a","b","c")', context),
        const TextValue('abc'),
      );
    });

    test('dependency graph tracks relationships', () {
      final graph = DependencyGraph();

      // B1 = A1 + 1, C1 = B1 * 2
      graph.updateDependencies('B1'.a1, {'A1'.a1});
      graph.updateDependencies('C1'.a1, {'B1'.a1});

      // Changing A1 should recalculate B1 then C1
      final toRecalc = graph.getCellsToRecalculate('A1'.a1);
      expect(toRecalc, ['B1'.a1, 'C1'.a1]);
    });

    test('custom function registration', () {
      final engine = FormulaEngine();
      engine.registerFunction(_DoubleFunction());

      final context = _SimpleContext(engine.functions);
      final result = engine.evaluateString('=DOUBLE(21)', context);
      expect(result, const NumberValue(42));
    });

    test('error handling', () {
      final engine = FormulaEngine();
      final context = _SimpleContext(engine.functions);

      // Division by zero
      final result = engine.evaluateString('=1/0', context);
      expect(result, const ErrorValue(FormulaError.divZero));

      // Unknown function
      final result2 = engine.evaluateString('=UNKNOWN(1)', context);
      expect(result2, const ErrorValue(FormulaError.name));
    });
  });
}

class _DoubleFunction extends FormulaFunction {
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
