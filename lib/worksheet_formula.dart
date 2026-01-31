/// A standalone formula engine for spreadsheet-like calculations.
///
/// This package provides:
/// - Excel/Google Sheets compatible formula parsing
/// - An extensible function registry with built-in functions
/// - Dependency tracking for efficient recalculation
/// - Type-safe formula values and error handling
///
/// ## Basic Usage
///
/// ```dart
/// import 'package:worksheet_formula/worksheet_formula.dart';
///
/// // Create the engine
/// final engine = FormulaEngine();
///
/// // Parse a formula
/// final ast = engine.parse('=SUM(A1:A10) * 2');
///
/// // Implement EvaluationContext to connect your data source
/// final context = MyEvaluationContext(myData);
///
/// // Evaluate
/// final result = engine.evaluate(ast, context);
/// ```
library;

// Core engine
export 'src/formula_engine.dart';

// AST
export 'src/ast/nodes.dart';
export 'src/ast/operators.dart';

// Evaluation
export 'src/evaluation/context.dart';
export 'src/evaluation/errors.dart';
export 'src/evaluation/value.dart';

// Functions
export 'src/functions/function.dart';
export 'src/functions/registry.dart';

// Dependencies
export 'src/dependencies/graph.dart';

// Parser (for advanced users)
export 'src/parser/formula_parser.dart';
