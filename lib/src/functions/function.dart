import '../ast/nodes.dart';
import '../evaluation/context.dart';
import '../evaluation/errors.dart';
import '../evaluation/value.dart';

/// Base class for all formula functions.
abstract class FormulaFunction {
  /// Function name (for error messages and registration).
  String get name;

  /// Minimum number of required arguments.
  int get minArgs;

  /// Maximum number of arguments (-1 for unlimited).
  int get maxArgs;

  /// Whether this function uses lazy evaluation (like IF).
  ///
  /// Lazy functions receive unevaluated FormulaNode arguments and
  /// can choose which ones to evaluate. Eager functions receive
  /// pre-evaluated FormulaValue arguments.
  bool get isLazy => false;

  /// Evaluate the function with the given arguments.
  FormulaValue call(List<FormulaNode> args, EvaluationContext context);

  /// Helper to evaluate all arguments eagerly.
  List<FormulaValue> evaluateArgs(
    List<FormulaNode> args,
    EvaluationContext context,
  ) {
    return args.map((arg) => arg.evaluate(context)).toList();
  }

  /// Helper to require a numeric value.
  FormulaValue requireNumber(FormulaValue value) {
    if (value.isError) return value;
    final n = value.toNumber();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    return FormulaValue.number(n);
  }

  /// Helper to collect all numbers from arguments (including ranges).
  Iterable<num> collectNumbers(
    List<FormulaNode> args,
    EvaluationContext context,
  ) sync* {
    for (final arg in args) {
      final value = arg.evaluate(context);
      switch (value) {
        case NumberValue(value: final n):
          yield n;
        case RangeValue(values: final matrix):
          for (final row in matrix) {
            for (final cell in row) {
              if (cell is NumberValue) {
                yield cell.value;
              }
            }
          }
        default:
          break;
      }
    }
  }
}
