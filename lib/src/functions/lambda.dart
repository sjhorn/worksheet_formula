import 'package:a1/a1.dart';

import '../ast/nodes.dart';
import '../evaluation/context.dart';
import '../evaluation/errors.dart';
import '../evaluation/value.dart';
import 'function.dart';
import 'registry.dart';

/// Register all lambda and higher-order functions.
void registerLambdaFunctions(FunctionRegistry registry) {
  registry.registerAll([
    LambdaFunction(),
    LetFunction(),
    MapFunction(),
    ReduceFunction(),
    ScanFunction(),
    MakeArrayFunction(),
    ByColFunction(),
    ByRowFunction(),
    IsOmittedFunction(),
  ]);
}

/// Evaluation context that layers variable bindings over a parent context.
class ScopedEvaluationContext implements EvaluationContext {
  final EvaluationContext _parent;
  final Map<String, FormulaValue> _variables;

  ScopedEvaluationContext(this._parent, Map<String, FormulaValue> variables)
      : _variables = {
          for (final entry in variables.entries)
            entry.key.toUpperCase(): entry.value,
        };

  @override
  FormulaValue getCellValue(A1 cell) => _parent.getCellValue(cell);

  @override
  FormulaValue getRangeValues(A1Range range) => _parent.getRangeValues(range);

  @override
  FormulaFunction? getFunction(String name) => _parent.getFunction(name);

  @override
  A1 get currentCell => _parent.currentCell;

  @override
  String? get currentSheet => _parent.currentSheet;

  @override
  bool get isCancelled => _parent.isCancelled;

  @override
  FormulaValue? getVariable(String name) {
    final value = _variables[name.toUpperCase()];
    if (value != null) return value;
    return _parent.getVariable(name);
  }
}

/// Invoke a FunctionValue with the given arguments.
///
/// Creates a ScopedEvaluationContext binding param names to arg values.
/// Missing args get OmittedValue.
FormulaValue invokeLambda(
  FunctionValue lambda,
  List<FormulaValue> args,
) {
  return lambda.invoke(args);
}

/// LAMBDA(param1, ..., paramN, body) — creates a reusable function.
///
/// All args except the last must be NameNodes (parameter names).
/// The last arg is the body expression. Returns a FunctionValue.
class LambdaFunction extends FormulaFunction {
  @override
  String get name => 'LAMBDA';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;
  @override
  bool get isLazy => true;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    if (args.length < 2) {
      // LAMBDA with only a body and no params is valid
      if (args.length == 1) {
        final body = args[0];
        return FunctionValue([], (callArgs) {
          return body.evaluate(context);
        });
      }
      return const FormulaValue.error(FormulaError.value);
    }

    // All args except the last must be NameNodes (param names)
    final paramNodes = args.sublist(0, args.length - 1);
    final paramNames = <String>[];
    for (final node in paramNodes) {
      if (node is! NameNode) {
        return const FormulaValue.error(FormulaError.value);
      }
      paramNames.add(node.name);
    }

    final body = args.last;

    return FunctionValue(paramNames, (callArgs) {
      final bindings = <String, FormulaValue>{};
      for (var i = 0; i < paramNames.length; i++) {
        bindings[paramNames[i]] =
            i < callArgs.length ? callArgs[i] : const OmittedValue();
      }
      final scopedContext = ScopedEvaluationContext(context, bindings);
      return body.evaluate(scopedContext);
    });
  }
}

/// LET(name1, value1, ..., nameN, valueN, expr) — binds names to values.
///
/// Must have an odd number of args >= 3. Names must be NameNodes.
/// Each binding can reference earlier bindings. The final expr is
/// evaluated in the scope of all bindings.
class LetFunction extends FormulaFunction {
  @override
  String get name => 'LET';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => -1;
  @override
  bool get isLazy => true;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    if (args.length % 2 == 0) {
      return const FormulaValue.error(FormulaError.value);
    }

    var currentContext = context;
    // Process name/value pairs
    for (var i = 0; i < args.length - 1; i += 2) {
      final nameNode = args[i];
      if (nameNode is! NameNode) {
        return const FormulaValue.error(FormulaError.value);
      }
      final value = args[i + 1].evaluate(currentContext);
      if (value.isError) return value;
      currentContext = ScopedEvaluationContext(
        currentContext,
        {nameNode.name: value},
      );
    }

    // Evaluate the final expression in the accumulated scope
    return args.last.evaluate(currentContext);
  }
}

/// MAP(array, lambda) — applies lambda to each element of array.
class MapFunction extends FormulaFunction {
  @override
  String get name => 'MAP';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;
  @override
  bool get isLazy => true;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final arrayVal = args[0].evaluate(context);
    final lambdaVal = args[1].evaluate(context);

    if (arrayVal.isError) return arrayVal;
    if (lambdaVal.isError) return lambdaVal;
    if (lambdaVal is! FunctionValue) {
      return const FormulaValue.error(FormulaError.value);
    }

    // Normalize to a 2D array
    final rows = _toRows(arrayVal);
    if (rows == null) return const FormulaValue.error(FormulaError.value);

    final result = <List<FormulaValue>>[];
    for (final row in rows) {
      final resultRow = <FormulaValue>[];
      for (final cell in row) {
        final value = invokeLambda(lambdaVal, [cell]);
        if (value.isError) return value;
        resultRow.add(value);
      }
      result.add(resultRow);
    }
    return FormulaValue.range(result);
  }
}

/// REDUCE(initial_value, array, lambda) — folds array with lambda(acc, elem).
class ReduceFunction extends FormulaFunction {
  @override
  String get name => 'REDUCE';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;
  @override
  bool get isLazy => true;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final initial = args[0].evaluate(context);
    if (initial.isError) return initial;

    final arrayVal = args[1].evaluate(context);
    if (arrayVal.isError) return arrayVal;

    final lambdaVal = args[2].evaluate(context);
    if (lambdaVal.isError) return lambdaVal;
    if (lambdaVal is! FunctionValue) {
      return const FormulaValue.error(FormulaError.value);
    }

    final rows = _toRows(arrayVal);
    if (rows == null) return const FormulaValue.error(FormulaError.value);

    var accumulator = initial;
    for (final row in rows) {
      for (final cell in row) {
        accumulator = invokeLambda(lambdaVal, [accumulator, cell]);
        if (accumulator.isError) return accumulator;
      }
    }
    return accumulator;
  }
}

/// SCAN(initial_value, array, lambda) — running fold, returns array of
/// intermediate accumulator values.
class ScanFunction extends FormulaFunction {
  @override
  String get name => 'SCAN';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;
  @override
  bool get isLazy => true;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final initial = args[0].evaluate(context);
    if (initial.isError) return initial;

    final arrayVal = args[1].evaluate(context);
    if (arrayVal.isError) return arrayVal;

    final lambdaVal = args[2].evaluate(context);
    if (lambdaVal.isError) return lambdaVal;
    if (lambdaVal is! FunctionValue) {
      return const FormulaValue.error(FormulaError.value);
    }

    final rows = _toRows(arrayVal);
    if (rows == null) return const FormulaValue.error(FormulaError.value);

    var accumulator = initial;
    final result = <List<FormulaValue>>[];
    for (final row in rows) {
      final resultRow = <FormulaValue>[];
      for (final cell in row) {
        accumulator = invokeLambda(lambdaVal, [accumulator, cell]);
        if (accumulator.isError) return accumulator;
        resultRow.add(accumulator);
      }
      result.add(resultRow);
    }
    return FormulaValue.range(result);
  }
}

/// MAKEARRAY(rows, cols, lambda) — builds array via lambda(row, col).
///
/// Row and col indices are 1-based (matching Excel behavior).
class MakeArrayFunction extends FormulaFunction {
  @override
  String get name => 'MAKEARRAY';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;
  @override
  bool get isLazy => true;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final rowsVal = args[0].evaluate(context);
    if (rowsVal.isError) return rowsVal;
    final colsVal = args[1].evaluate(context);
    if (colsVal.isError) return colsVal;

    final rowCount = rowsVal.toNumber();
    final colCount = colsVal.toNumber();
    if (rowCount == null || colCount == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    final rows = rowCount.toInt();
    final cols = colCount.toInt();
    if (rows < 1 || cols < 1) {
      return const FormulaValue.error(FormulaError.value);
    }

    final lambdaVal = args[2].evaluate(context);
    if (lambdaVal.isError) return lambdaVal;
    if (lambdaVal is! FunctionValue) {
      return const FormulaValue.error(FormulaError.value);
    }

    final result = <List<FormulaValue>>[];
    for (var r = 1; r <= rows; r++) {
      final resultRow = <FormulaValue>[];
      for (var c = 1; c <= cols; c++) {
        final value = invokeLambda(
          lambdaVal,
          [FormulaValue.number(r), FormulaValue.number(c)],
        );
        if (value.isError) return value;
        resultRow.add(value);
      }
      result.add(resultRow);
    }
    return FormulaValue.range(result);
  }
}

/// BYCOL(array, lambda) — applies lambda to each column of the array.
///
/// Lambda receives a single-column RangeValue for each column.
class ByColFunction extends FormulaFunction {
  @override
  String get name => 'BYCOL';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;
  @override
  bool get isLazy => true;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final arrayVal = args[0].evaluate(context);
    if (arrayVal.isError) return arrayVal;

    final lambdaVal = args[1].evaluate(context);
    if (lambdaVal.isError) return lambdaVal;
    if (lambdaVal is! FunctionValue) {
      return const FormulaValue.error(FormulaError.value);
    }

    final rows = _toRows(arrayVal);
    if (rows == null || rows.isEmpty) {
      return const FormulaValue.error(FormulaError.value);
    }

    final colCount = rows.first.length;
    final resultRow = <FormulaValue>[];
    for (var c = 0; c < colCount; c++) {
      // Extract column as a single-column RangeValue
      final column = <List<FormulaValue>>[];
      for (final row in rows) {
        column.add([row[c]]);
      }
      final colRange = FormulaValue.range(column);
      final value = invokeLambda(lambdaVal, [colRange]);
      if (value.isError) return value;
      resultRow.add(value);
    }
    return FormulaValue.range([resultRow]);
  }
}

/// BYROW(array, lambda) — applies lambda to each row of the array.
///
/// Lambda receives a single-row RangeValue for each row.
class ByRowFunction extends FormulaFunction {
  @override
  String get name => 'BYROW';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;
  @override
  bool get isLazy => true;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final arrayVal = args[0].evaluate(context);
    if (arrayVal.isError) return arrayVal;

    final lambdaVal = args[1].evaluate(context);
    if (lambdaVal.isError) return lambdaVal;
    if (lambdaVal is! FunctionValue) {
      return const FormulaValue.error(FormulaError.value);
    }

    final rows = _toRows(arrayVal);
    if (rows == null || rows.isEmpty) {
      return const FormulaValue.error(FormulaError.value);
    }

    final resultCol = <List<FormulaValue>>[];
    for (final row in rows) {
      final rowRange = FormulaValue.range([row]);
      final value = invokeLambda(lambdaVal, [rowRange]);
      if (value.isError) return value;
      resultCol.add([value]);
    }
    return FormulaValue.range(resultCol);
  }
}

/// ISOMITTED(value) — returns TRUE if value is an OmittedValue sentinel.
class IsOmittedFunction extends FormulaFunction {
  @override
  String get name => 'ISOMITTED';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    return FormulaValue.boolean(value is OmittedValue);
  }
}

/// Helper: normalize a FormulaValue to a 2D list of values.
///
/// Single values become a 1x1 array. RangeValues return their rows.
List<List<FormulaValue>>? _toRows(FormulaValue value) {
  if (value is RangeValue) {
    return value.values;
  }
  // Single value → 1x1 array
  return [
    [value]
  ];
}
