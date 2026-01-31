import 'package:a1/a1.dart';
import 'package:test/test.dart';
import 'package:worksheet_formula/src/ast/nodes.dart';
import 'package:worksheet_formula/src/evaluation/context.dart';
import 'package:worksheet_formula/src/evaluation/errors.dart';
import 'package:worksheet_formula/src/evaluation/value.dart';
import 'package:worksheet_formula/src/formula_engine.dart';
import 'package:worksheet_formula/src/functions/function.dart';
import 'package:worksheet_formula/src/parser/formula_parser.dart';
import 'package:worksheet_formula/src/functions/registry.dart';

class _TestContext implements EvaluationContext {
  final Map<A1, FormulaValue> cells;
  final FunctionRegistry _registry;

  _TestContext(this._registry, [Map<A1, FormulaValue>? cells])
      : cells = cells ?? {};

  @override
  A1 get currentCell => 'A1'.a1;
  @override
  String? get currentSheet => null;
  @override
  bool get isCancelled => false;

  @override
  FormulaValue getCellValue(A1 cell) =>
      cells[cell] ?? const EmptyValue();

  @override
  FormulaValue getRangeValues(A1Range range) =>
      const FormulaValue.error(FormulaError.ref);

  @override
  FormulaFunction? getFunction(String name) => _registry.get(name);
}

void main() {
  late FormulaEngine engine;

  setUp(() {
    engine = FormulaEngine();
  });

  group('parse', () {
    test('parses valid formula', () {
      final ast = engine.parse('=1+2');
      expect(ast, isA<BinaryOpNode>());
    });

    test('caches parsed results', () {
      final ast1 = engine.parse('=1+2');
      final ast2 = engine.parse('=1+2');
      expect(identical(ast1, ast2), isTrue);
    });

    test('throws on invalid formula', () {
      expect(() => engine.parse('=+'), throwsA(isA<FormulaParseException>()));
    });
  });

  group('tryParse', () {
    test('returns node on success', () {
      expect(engine.tryParse('=1+2'), isNotNull);
    });

    test('returns null on failure', () {
      expect(engine.tryParse('=+'), isNull);
    });
  });

  group('isValidFormula', () {
    test('returns true for valid formulas', () {
      expect(engine.isValidFormula('=1+2'), isTrue);
      expect(engine.isValidFormula('=SUM(1,2)'), isTrue);
    });

    test('returns false for invalid formulas', () {
      expect(engine.isValidFormula('=+'), isFalse);
    });
  });

  group('evaluate', () {
    test('evaluates a parsed AST', () {
      final ast = engine.parse('=1+2');
      final context = _TestContext(engine.functions);
      final result = engine.evaluate(ast, context);
      expect(result, const NumberValue(3));
    });
  });

  group('evaluateString', () {
    test('parses and evaluates in one call', () {
      final context = _TestContext(engine.functions);
      final result = engine.evaluateString('=2*3', context);
      expect(result, const NumberValue(6));
    });

    test('evaluates formula with cell references', () {
      final context = _TestContext(
        engine.functions,
        {'A1'.a1: const NumberValue(10)},
      );
      final result = engine.evaluateString('=A1+5', context);
      expect(result, const NumberValue(15));
    });

    test('evaluates function calls', () {
      final context = _TestContext(engine.functions);
      final result = engine.evaluateString('=SUM(1,2,3)', context);
      expect(result, const NumberValue(6));
    });

    test('evaluates IF function', () {
      final context = _TestContext(engine.functions);
      final result =
          engine.evaluateString('=IF(TRUE,"yes","no")', context);
      expect(result, const TextValue('yes'));
    });

    test('evaluates text functions', () {
      final context = _TestContext(engine.functions);
      final result = engine.evaluateString('=LEN("hello")', context);
      expect(result, const NumberValue(5));
    });
  });

  group('getCellReferences', () {
    test('returns cell references from formula', () {
      final refs = engine.getCellReferences('=A1+B2');
      expect(refs, containsAll(['A1'.a1, 'B2'.a1]));
    });

    test('returns empty set for formula with no refs', () {
      final refs = engine.getCellReferences('=1+2');
      expect(refs, isEmpty);
    });

    test('returns references from function arguments', () {
      final refs = engine.getCellReferences('=SUM(A1,B1)');
      expect(refs, containsAll(['A1'.a1, 'B1'.a1]));
    });
  });

  group('registerFunction', () {
    test('registers custom function', () {
      engine.registerFunction(_DoubleFunction());
      final context = _TestContext(engine.functions);
      final result = engine.evaluateString('=DOUBLE(21)', context);
      expect(result, const NumberValue(42));
    });
  });

  group('clearCache', () {
    test('clears parse cache', () {
      final ast1 = engine.parse('=1+2');
      engine.clearCache();
      final ast2 = engine.parse('=1+2');
      // After clearing cache, should be a different object
      expect(identical(ast1, ast2), isFalse);
    });
  });

  group('functions', () {
    test('exposes function registry', () {
      expect(engine.functions, isA<FunctionRegistry>());
      expect(engine.functions.has('SUM'), isTrue);
      expect(engine.functions.has('IF'), isTrue);
      expect(engine.functions.has('CONCAT'), isTrue);
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
