## 0.2.0

- **Statistical functions**: COUNT, COUNTA, COUNTBLANK, COUNTIF, SUMIF, AVERAGEIF with shared criteria matching
- **Lookup functions**: VLOOKUP (exact & approximate), INDEX, MATCH (exact, ascending, descending)
- **Date functions**: DATE, TODAY, NOW, YEAR, MONTH, DAY using Excel serial number convention
- **TEXT format codes**: `0`, `#`, decimal, thousands separator, percentage, scientific notation
- **Better parser errors**: Position-aware messages for unmatched parentheses, unexpected tokens, truncated input
- **Performance benchmarks**: Parse, evaluate, and dependency graph benchmarks
- **Iterative graph traversal**: DependencyGraph uses iterative DFS to handle deep cell chains
- **Worksheet example**: Runnable example demonstrating full API usage
- **Flutter integration example**: Standalone Flutter app integrating with the `worksheet` widget
- **MIT license**

## 0.1.0

- Initial version with core formula engine
- Formula parser with operator precedence (arithmetic, comparison, concatenation, percent, power)
- AST representation with sealed FormulaNode classes
- FormulaValue type system: Number, Text, Boolean, Error, Empty, Range
- Excel-compatible error types: #DIV/0!, #VALUE!, #REF!, #NAME?, #NUM!, #N/A, #NULL!, #CALC!, #CIRCULAR!
- Math functions: SUM, AVERAGE, MIN, MAX, ABS, ROUND, INT, MOD, SQRT, POWER
- Logical functions: IF, AND, OR, NOT, IFERROR, IFNA, TRUE, FALSE
- Text functions: CONCAT, CONCATENATE, LEFT, RIGHT, MID, LEN, LOWER, UPPER, TRIM, TEXT
- EvaluationContext interface for pluggable data sources
- FunctionRegistry with custom function registration
- DependencyGraph for cell dependency tracking and recalculation ordering
- Parse caching for performance
- Cell reference extraction from formulas
