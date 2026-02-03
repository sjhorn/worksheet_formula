import 'package:a1/a1.dart';

import '../functions/function.dart';
import 'value.dart';

/// Abstract context for formula evaluation.
///
/// Implement this interface to connect the formula engine to your data source.
/// This is the primary integration point for consumers of the package.
abstract class EvaluationContext {
  /// Get the value of a single cell.
  FormulaValue getCellValue(A1 cell);

  /// Get values for a range of cells as a 2D matrix.
  FormulaValue getRangeValues(A1Range range);

  /// Get a function by name (case-insensitive).
  FormulaFunction? getFunction(String name);

  /// The cell currently being evaluated (for relative references).
  A1 get currentCell;

  /// The current sheet name (for cross-sheet references).
  String? get currentSheet;

  /// Get a variable value by name (for LAMBDA parameter scoping).
  ///
  /// Returns null if the variable is not defined in this scope.
  FormulaValue? getVariable(String name) => null;

  /// Optional: Check if evaluation should be cancelled (for long-running calcs).
  bool get isCancelled => false;
}
