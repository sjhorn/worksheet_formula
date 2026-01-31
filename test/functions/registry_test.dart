import 'package:test/test.dart';
import 'package:worksheet_formula/src/ast/nodes.dart';
import 'package:worksheet_formula/src/evaluation/context.dart';
import 'package:worksheet_formula/src/evaluation/value.dart';
import 'package:worksheet_formula/src/functions/function.dart';
import 'package:worksheet_formula/src/functions/registry.dart';

class _StubFunction extends FormulaFunction {
  @override
  final String name;
  @override
  int get minArgs => 0;
  @override
  int get maxArgs => -1;
  _StubFunction(this.name);
  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) =>
      const NumberValue(0);
}

void main() {
  group('FunctionRegistry', () {
    test('register and get function', () {
      final registry = FunctionRegistry(registerBuiltIns: false);
      registry.register(_StubFunction('FOO'));
      expect(registry.get('FOO'), isNotNull);
    });

    test('get is case-insensitive', () {
      final registry = FunctionRegistry(registerBuiltIns: false);
      registry.register(_StubFunction('FOO'));
      expect(registry.get('foo'), isNotNull);
      expect(registry.get('Foo'), isNotNull);
    });

    test('get returns null for unknown function', () {
      final registry = FunctionRegistry(registerBuiltIns: false);
      expect(registry.get('UNKNOWN'), isNull);
    });

    test('has checks existence', () {
      final registry = FunctionRegistry(registerBuiltIns: false);
      registry.register(_StubFunction('FOO'));
      expect(registry.has('FOO'), isTrue);
      expect(registry.has('BAR'), isFalse);
    });

    test('names returns registered names', () {
      final registry = FunctionRegistry(registerBuiltIns: false);
      registry.register(_StubFunction('FOO'));
      registry.register(_StubFunction('BAR'));
      expect(registry.names, containsAll(['FOO', 'BAR']));
    });

    test('registerAll adds multiple functions', () {
      final registry = FunctionRegistry(registerBuiltIns: false);
      registry.registerAll([_StubFunction('A'), _StubFunction('B')]);
      expect(registry.has('A'), isTrue);
      expect(registry.has('B'), isTrue);
    });

    test('copyWith creates copy with additional functions', () {
      final registry = FunctionRegistry(registerBuiltIns: false);
      registry.register(_StubFunction('FOO'));
      final copy = registry.copyWith([_StubFunction('BAR')]);
      expect(copy.has('FOO'), isTrue);
      expect(copy.has('BAR'), isTrue);
      // Original unchanged
      expect(registry.has('BAR'), isFalse);
    });

    test('registerBuiltIns true registers math functions', () {
      final registry = FunctionRegistry();
      expect(registry.has('SUM'), isTrue);
      expect(registry.has('AVERAGE'), isTrue);
      expect(registry.has('MIN'), isTrue);
      expect(registry.has('MAX'), isTrue);
    });
  });
}
