## 0.5.0

- **Financial functions** (40): PMT, FV, PV, NPER, RATE, IPMT, PPMT, CUMIPMT, CUMPRINC, NPV, XNPV, IRR, XIRR, MIRR, FVSCHEDULE, SLN, SYD, DB, DDB, VDB, PRICE, YIELD, DURATION, MDURATION, ACCRINT, DISC, INTRATE, RECEIVED, PRICEDISC, PRICEMAT, TBILLEQ, TBILLPRICE, TBILLYIELD, DOLLARDE, DOLLARFR, EFFECT, NOMINAL, PDURATION, RRI, ISPMT
- Loans & annuities (TVM): PMT, FV, PV, NPER, RATE, IPMT, PPMT, CUMIPMT, CUMPRINC
- Investment analysis: NPV, XNPV, IRR, XIRR, MIRR, FVSCHEDULE
- Depreciation: SLN, SYD, DB, DDB, VDB
- Bonds & securities: PRICE, YIELD, DURATION, MDURATION, ACCRINT, DISC, INTRATE, RECEIVED, PRICEDISC, PRICEMAT, TBILLEQ, TBILLPRICE, TBILLYIELD, DOLLARDE, DOLLARFR
- Other financial: EFFECT, NOMINAL, PDURATION, RRI, ISPMT
- Newton-Raphson solver for RATE, IRR, XIRR, YIELD
- Day count basis support (US 30/360, actual/actual, actual/360, actual/365, EU 30/360)
- 239 total built-in functions across 9 categories

## 0.4.0

- **Dynamic array functions** (17): SEQUENCE, RANDARRAY, TOCOL, TOROW, WRAPROWS, WRAPCOLS, CHOOSEROWS, CHOOSECOLS, DROP, TAKE, EXPAND, HSTACK, VSTACK, FILTER, UNIQUE, SORT, SORTBY
- Array generators, flatten/reshape, slice/select, concatenation, and filter/sort/unique operations
- Private shared helpers for matrix conversion, flattening, value comparison, and index resolution
- 199 total built-in functions across 8 categories

## 0.3.0

- **87 new built-in functions** across all existing categories (95 → 182 total)
- **Math & Trigonometry** (28 new): PI, LN, LOG, LOG10, EXP, SIN, COS, TAN, ASIN, ACOS, ATAN, ATAN2, DEGREES, RADIANS, EVEN, ODD, GCD, LCM, TRUNC, MROUND, QUOTIENT, COMBIN, COMBINA, FACT, FACTDOUBLE, SUMSQ, SUBTOTAL, AGGREGATE
- **Text** (13 new): REPT, CHAR, CODE, CLEAN, DOLLAR, FIXED, T, NUMBERVALUE, UNICHAR, UNICODE, TEXTBEFORE, TEXTAFTER, TEXTSPLIT
- **Statistical** (19 new): STDEV.S, STDEV.P, VAR.S, VAR.P, PERCENTILE.INC, PERCENTILE.EXC, PERCENTRANK.INC, PERCENTRANK.EXC, RANK.AVG, FREQUENCY, AVEDEV, AVERAGEA, MAXA, MINA, TRIMMEAN, GEOMEAN, HARMEAN, MAXIFS, MINIFS
- **Lookup & Reference** (10 new): ROW, COLUMN, ROWS, COLUMNS, ADDRESS, INDIRECT, OFFSET, TRANSPOSE, HYPERLINK, AREAS
- **Date/Time** (9 new): TIMEVALUE, WEEKNUM, ISOWEEKNUM, NETWORKDAYS, NETWORKDAYS.INTL, WORKDAY, WORKDAY.INTL, DAYS360, YEARFRAC
- **Information** (8 new): ISERR, ISNONTEXT, ISEVEN, ISODD, ISREF, N, NA, ERROR.TYPE

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
