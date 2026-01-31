import 'package:a1/a1.dart';

import 'ast/nodes.dart';
import 'evaluation/context.dart';
import 'evaluation/value.dart';
import 'functions/function.dart';
import 'functions/registry.dart';
import 'parser/formula_parser.dart';

/// Main entry point for the formula engine.
///
/// Use this class to parse and evaluate spreadsheet formulas.
///
/// ```dart
/// final engine = FormulaEngine();
///
/// // Parse a formula
/// final ast = engine.parse('=SUM(A1:A10)');
///
/// // Evaluate with your data source
/// final result = engine.evaluate(ast, myContext);
/// ```
class FormulaEngine {
  final FormulaParser _parser;
  final FunctionRegistry _functions;
  final Map<String, FormulaNode> _parseCache = {};

  /// Create a new formula engine.
  ///
  /// Optionally provide a custom [FunctionRegistry] to use different
  /// or additional functions.
  FormulaEngine({FunctionRegistry? functions})
      : _parser = FormulaParser(),
        _functions = functions ?? FunctionRegistry();

  /// Parse a formula string into an AST.
  ///
  /// Results are cached for performance. Throws [FormulaParseException]
  /// if the formula is invalid.
  FormulaNode parse(String formula) {
    return _parseCache.putIfAbsent(
      formula,
      () => _parser.parse(formula),
    );
  }

  /// Try to parse a formula, returning null on failure.
  FormulaNode? tryParse(String formula) {
    try {
      return parse(formula);
    } catch (_) {
      return null;
    }
  }

  /// Check if a string is a valid formula.
  bool isValidFormula(String formula) => tryParse(formula) != null;

  /// Evaluate a parsed formula with the given context.
  FormulaValue evaluate(FormulaNode ast, EvaluationContext context) {
    return ast.evaluate(context);
  }

  /// Parse and evaluate a formula string.
  FormulaValue evaluateString(String formula, EvaluationContext context) {
    final ast = parse(formula);
    return evaluate(ast, context);
  }

  /// Get all cell references in a formula (for dependency tracking).
  Set<A1> getCellReferences(String formula) {
    final ast = parse(formula);
    return ast.cellReferences.toSet();
  }

  /// Register a custom function.
  void registerFunction(FormulaFunction function) {
    _functions.register(function);
  }

  /// Get the function registry.
  ///
  /// Use this when creating evaluation contexts.
  FunctionRegistry get functions => _functions;

  /// Clear the parse cache.
  void clearCache() => _parseCache.clear();
}
