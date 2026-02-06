# worksheet_formula

A standalone formula engine for spreadsheet-like calculations in Dart.

## Features

- Excel/Google Sheets compatible formula parsing
- 400 built-in functions across 14 categories (math, logical, text, statistical, statistical advanced, lookup, date, information, array, financial, engineering, database, lambda/higher-order, web/regex)
- Dynamic array functions (FILTER, SORT, UNIQUE, SEQUENCE, etc.)
- Type-safe formula values with Excel-compatible error handling
- Cell dependency tracking for efficient recalculation
- Custom function registration
- Parse caching for performance
- Zero UI dependencies -- works with any data source

## Installation

```yaml
dependencies:
  worksheet_formula: ^1.0.0
```

## Quick Start

```dart
import 'package:worksheet_formula/worksheet_formula.dart';

// Create the engine
final engine = FormulaEngine();

// Parse and evaluate
final ast = engine.parse('=1+2*3');
final result = engine.evaluate(ast, myContext);
// result = NumberValue(7)
```

To connect your data, implement `EvaluationContext`:

```dart
class MyContext implements EvaluationContext {
  final Map<A1, FormulaValue> cells;
  final FunctionRegistry registry;

  MyContext(this.registry, this.cells);

  @override
  A1 get currentCell => 'A1'.a1;
  @override
  String? get currentSheet => null;
  @override
  bool get isCancelled => false;

  @override
  FormulaValue getCellValue(A1 cell) => cells[cell] ?? const EmptyValue();

  @override
  FormulaValue getRangeValues(A1Range range) {
    // Build 2D matrix from your data source
    // ...
  }

  @override
  FormulaFunction? getFunction(String name) => registry.get(name);

  @override
  FormulaValue? getVariable(String name) => null;
}
```

Then evaluate formulas with cell references:

```dart
final context = MyContext(engine.functions, {
  'A1'.a1: NumberValue(10),
  'A2'.a1: NumberValue(20),
});
final result = engine.evaluateString('=SUM(A1:A2)', context);
// result = NumberValue(30)
```

## Built-in Functions

### [Math & Trigonometry (50)](doc/functions/math.md)
[`SUM`](doc/functions/math.md#sum), [`AVERAGE`](doc/functions/math.md#average), [`MIN`](doc/functions/math.md#min), [`MAX`](doc/functions/math.md#max), [`ABS`](doc/functions/math.md#abs), [`ROUND`](doc/functions/math.md#round), [`INT`](doc/functions/math.md#int), [`MOD`](doc/functions/math.md#mod), [`SQRT`](doc/functions/math.md#sqrt), [`POWER`](doc/functions/math.md#power), [`SUMPRODUCT`](doc/functions/math.md#sumproduct), [`ROUNDUP`](doc/functions/math.md#roundup), [`ROUNDDOWN`](doc/functions/math.md#rounddown), [`CEILING`](doc/functions/math.md#ceiling), [`FLOOR`](doc/functions/math.md#floor), [`SIGN`](doc/functions/math.md#sign), [`PRODUCT`](doc/functions/math.md#product), [`RAND`](doc/functions/math.md#rand), [`RANDBETWEEN`](doc/functions/math.md#randbetween), [`PI`](doc/functions/math.md#pi), [`LN`](doc/functions/math.md#ln), [`LOG`](doc/functions/math.md#log), [`LOG10`](doc/functions/math.md#log10), [`EXP`](doc/functions/math.md#exp), [`SIN`](doc/functions/math.md#sin), [`COS`](doc/functions/math.md#cos), [`TAN`](doc/functions/math.md#tan), [`ASIN`](doc/functions/math.md#asin), [`ACOS`](doc/functions/math.md#acos), [`ATAN`](doc/functions/math.md#atan), [`ATAN2`](doc/functions/math.md#atan2), [`DEGREES`](doc/functions/math.md#degrees), [`RADIANS`](doc/functions/math.md#radians), [`EVEN`](doc/functions/math.md#even), [`ODD`](doc/functions/math.md#odd), [`GCD`](doc/functions/math.md#gcd), [`LCM`](doc/functions/math.md#lcm), [`TRUNC`](doc/functions/math.md#trunc), [`MROUND`](doc/functions/math.md#mround), [`QUOTIENT`](doc/functions/math.md#quotient), [`COMBIN`](doc/functions/math.md#combin), [`COMBINA`](doc/functions/math.md#combina), [`FACT`](doc/functions/math.md#fact), [`FACTDOUBLE`](doc/functions/math.md#factdouble), [`SUMSQ`](doc/functions/math.md#sumsq), [`SUBTOTAL`](doc/functions/math.md#subtotal), [`AGGREGATE`](doc/functions/math.md#aggregate), [`SERIESSUM`](doc/functions/math.md#seriessum), [`SQRTPI`](doc/functions/math.md#sqrtpi), [`MULTINOMIAL`](doc/functions/math.md#multinomial)

### [Logical (11)](doc/functions/logical.md)
[`IF`](doc/functions/logical.md#if), [`AND`](doc/functions/logical.md#and), [`OR`](doc/functions/logical.md#or), [`NOT`](doc/functions/logical.md#not), [`IFERROR`](doc/functions/logical.md#iferror), [`IFNA`](doc/functions/logical.md#ifna), [`TRUE`](doc/functions/logical.md#true), [`FALSE`](doc/functions/logical.md#false), [`IFS`](doc/functions/logical.md#ifs), [`SWITCH`](doc/functions/logical.md#switch), [`XOR`](doc/functions/logical.md#xor)

### [Text (36)](doc/functions/text.md)
[`CONCAT`](doc/functions/text.md#concat), [`CONCATENATE`](doc/functions/text.md#concatenate), [`LEFT`](doc/functions/text.md#left), [`RIGHT`](doc/functions/text.md#right), [`MID`](doc/functions/text.md#mid), [`LEN`](doc/functions/text.md#len), [`LOWER`](doc/functions/text.md#lower), [`UPPER`](doc/functions/text.md#upper), [`TRIM`](doc/functions/text.md#trim), [`TEXT`](doc/functions/text.md#text), [`FIND`](doc/functions/text.md#find), [`SEARCH`](doc/functions/text.md#search), [`SUBSTITUTE`](doc/functions/text.md#substitute), [`REPLACE`](doc/functions/text.md#replace), [`VALUE`](doc/functions/text.md#value), [`TEXTJOIN`](doc/functions/text.md#textjoin), [`PROPER`](doc/functions/text.md#proper), [`EXACT`](doc/functions/text.md#exact), [`REPT`](doc/functions/text.md#rept), [`CHAR`](doc/functions/text.md#char), [`CODE`](doc/functions/text.md#code), [`CLEAN`](doc/functions/text.md#clean), [`DOLLAR`](doc/functions/text.md#dollar), [`FIXED`](doc/functions/text.md#fixed), [`T`](doc/functions/text.md#t), [`NUMBERVALUE`](doc/functions/text.md#numbervalue), [`UNICHAR`](doc/functions/text.md#unichar), [`UNICODE`](doc/functions/text.md#unicode), [`TEXTBEFORE`](doc/functions/text.md#textbefore), [`TEXTAFTER`](doc/functions/text.md#textafter), [`TEXTSPLIT`](doc/functions/text.md#textsplit), [`ARRAYTOTEXT`](doc/functions/text.md#arraytotext), [`VALUETOTEXT`](doc/functions/text.md#valuetotext), [`ASC`](doc/functions/text.md#asc), [`DBCS`](doc/functions/text.md#dbcs), [`BAHTTEXT`](doc/functions/text.md#bahttext)

### [Statistical (35)](doc/functions/statistical.md)
[`COUNT`](doc/functions/statistical.md#count), [`COUNTA`](doc/functions/statistical.md#counta), [`COUNTBLANK`](doc/functions/statistical.md#countblank), [`COUNTIF`](doc/functions/statistical.md#countif), [`SUMIF`](doc/functions/statistical.md#sumif), [`AVERAGEIF`](doc/functions/statistical.md#averageif), [`SUMIFS`](doc/functions/statistical.md#sumifs), [`COUNTIFS`](doc/functions/statistical.md#countifs), [`AVERAGEIFS`](doc/functions/statistical.md#averageifs), [`MEDIAN`](doc/functions/statistical.md#median), [`MODE.SNGL`](doc/functions/statistical.md#modesngl), [`MODE`](doc/functions/statistical.md#mode), [`LARGE`](doc/functions/statistical.md#large), [`SMALL`](doc/functions/statistical.md#small), [`RANK.EQ`](doc/functions/statistical.md#rankeq), [`RANK`](doc/functions/statistical.md#rank), [`STDEV.S`](doc/functions/statistical.md#stdevs), [`STDEV.P`](doc/functions/statistical.md#stdevp), [`VAR.S`](doc/functions/statistical.md#vars), [`VAR.P`](doc/functions/statistical.md#varp), [`PERCENTILE.INC`](doc/functions/statistical.md#percentileinc), [`PERCENTILE.EXC`](doc/functions/statistical.md#percentileexc), [`PERCENTRANK.INC`](doc/functions/statistical.md#percentrankinc), [`PERCENTRANK.EXC`](doc/functions/statistical.md#percentrankexc), [`RANK.AVG`](doc/functions/statistical.md#rankavg), [`FREQUENCY`](doc/functions/statistical.md#frequency), [`AVEDEV`](doc/functions/statistical.md#avedev), [`AVERAGEA`](doc/functions/statistical.md#averagea), [`MAXA`](doc/functions/statistical.md#maxa), [`MINA`](doc/functions/statistical.md#mina), [`TRIMMEAN`](doc/functions/statistical.md#trimmean), [`GEOMEAN`](doc/functions/statistical.md#geomean), [`HARMEAN`](doc/functions/statistical.md#harmean), [`MAXIFS`](doc/functions/statistical.md#maxifs), [`MINIFS`](doc/functions/statistical.md#minifs)

### [Lookup & Reference (18)](doc/functions/lookup.md)
[`VLOOKUP`](doc/functions/lookup.md#vlookup), [`INDEX`](doc/functions/lookup.md#index), [`MATCH`](doc/functions/lookup.md#match), [`HLOOKUP`](doc/functions/lookup.md#hlookup), [`LOOKUP`](doc/functions/lookup.md#lookup), [`CHOOSE`](doc/functions/lookup.md#choose), [`XMATCH`](doc/functions/lookup.md#xmatch), [`XLOOKUP`](doc/functions/lookup.md#xlookup), [`ROW`](doc/functions/lookup.md#row), [`COLUMN`](doc/functions/lookup.md#column), [`ROWS`](doc/functions/lookup.md#rows), [`COLUMNS`](doc/functions/lookup.md#columns), [`ADDRESS`](doc/functions/lookup.md#address), [`INDIRECT`](doc/functions/lookup.md#indirect), [`OFFSET`](doc/functions/lookup.md#offset), [`TRANSPOSE`](doc/functions/lookup.md#transpose), [`HYPERLINK`](doc/functions/lookup.md#hyperlink), [`AREAS`](doc/functions/lookup.md#areas)

### [Date/Time (25)](doc/functions/date.md)
[`DATE`](doc/functions/date.md#date), [`TODAY`](doc/functions/date.md#today), [`NOW`](doc/functions/date.md#now), [`YEAR`](doc/functions/date.md#year), [`MONTH`](doc/functions/date.md#month), [`DAY`](doc/functions/date.md#day), [`DAYS`](doc/functions/date.md#days), [`DATEDIF`](doc/functions/date.md#datedif), [`DATEVALUE`](doc/functions/date.md#datevalue), [`WEEKDAY`](doc/functions/date.md#weekday), [`HOUR`](doc/functions/date.md#hour), [`MINUTE`](doc/functions/date.md#minute), [`SECOND`](doc/functions/date.md#second), [`TIME`](doc/functions/date.md#time), [`EDATE`](doc/functions/date.md#edate), [`EOMONTH`](doc/functions/date.md#eomonth), [`TIMEVALUE`](doc/functions/date.md#timevalue), [`WEEKNUM`](doc/functions/date.md#weeknum), [`ISOWEEKNUM`](doc/functions/date.md#isoweeknum), [`NETWORKDAYS`](doc/functions/date.md#networkdays), [`NETWORKDAYS.INTL`](doc/functions/date.md#networkdaysintl), [`WORKDAY`](doc/functions/date.md#workday), [`WORKDAY.INTL`](doc/functions/date.md#workdayintl), [`DAYS360`](doc/functions/date.md#days360), [`YEARFRAC`](doc/functions/date.md#yearfrac)

### [Information (15)](doc/functions/information.md)
[`ISBLANK`](doc/functions/information.md#isblank), [`ISERROR`](doc/functions/information.md#iserror), [`ISNUMBER`](doc/functions/information.md#isnumber), [`ISTEXT`](doc/functions/information.md#istext), [`ISLOGICAL`](doc/functions/information.md#islogical), [`ISNA`](doc/functions/information.md#isna), [`TYPE`](doc/functions/information.md#type), [`ISERR`](doc/functions/information.md#iserr), [`ISNONTEXT`](doc/functions/information.md#isnontext), [`ISEVEN`](doc/functions/information.md#iseven), [`ISODD`](doc/functions/information.md#isodd), [`ISREF`](doc/functions/information.md#isref), [`N`](doc/functions/information.md#n), [`NA`](doc/functions/information.md#na), [`ERROR.TYPE`](doc/functions/information.md#errortype)

### [Dynamic Array (21)](doc/functions/array.md)
[`SEQUENCE`](doc/functions/array.md#sequence), [`RANDARRAY`](doc/functions/array.md#randarray), [`TOCOL`](doc/functions/array.md#tocol), [`TOROW`](doc/functions/array.md#torow), [`WRAPROWS`](doc/functions/array.md#wraprows), [`WRAPCOLS`](doc/functions/array.md#wrapcols), [`CHOOSEROWS`](doc/functions/array.md#chooserows), [`CHOOSECOLS`](doc/functions/array.md#choosecols), [`DROP`](doc/functions/array.md#drop), [`TAKE`](doc/functions/array.md#take), [`EXPAND`](doc/functions/array.md#expand), [`HSTACK`](doc/functions/array.md#hstack), [`VSTACK`](doc/functions/array.md#vstack), [`FILTER`](doc/functions/array.md#filter), [`UNIQUE`](doc/functions/array.md#unique), [`SORT`](doc/functions/array.md#sort), [`SORTBY`](doc/functions/array.md#sortby), [`MUNIT`](doc/functions/array.md#munit), [`MMULT`](doc/functions/array.md#mmult), [`MDETERM`](doc/functions/array.md#mdeterm), [`MINVERSE`](doc/functions/array.md#minverse)

### [Financial (40)](doc/functions/financial.md)
[`PMT`](doc/functions/financial.md#pmt), [`FV`](doc/functions/financial.md#fv), [`PV`](doc/functions/financial.md#pv), [`NPER`](doc/functions/financial.md#nper), [`RATE`](doc/functions/financial.md#rate), [`IPMT`](doc/functions/financial.md#ipmt), [`PPMT`](doc/functions/financial.md#ppmt), [`CUMIPMT`](doc/functions/financial.md#cumipmt), [`CUMPRINC`](doc/functions/financial.md#cumprinc), [`NPV`](doc/functions/financial.md#npv), [`XNPV`](doc/functions/financial.md#xnpv), [`IRR`](doc/functions/financial.md#irr), [`XIRR`](doc/functions/financial.md#xirr), [`MIRR`](doc/functions/financial.md#mirr), [`FVSCHEDULE`](doc/functions/financial.md#fvschedule), [`SLN`](doc/functions/financial.md#sln), [`SYD`](doc/functions/financial.md#syd), [`DB`](doc/functions/financial.md#db), [`DDB`](doc/functions/financial.md#ddb), [`VDB`](doc/functions/financial.md#vdb), [`PRICE`](doc/functions/financial.md#price), [`YIELD`](doc/functions/financial.md#yield), [`DURATION`](doc/functions/financial.md#duration), [`MDURATION`](doc/functions/financial.md#mduration), [`ACCRINT`](doc/functions/financial.md#accrint), [`DISC`](doc/functions/financial.md#disc), [`INTRATE`](doc/functions/financial.md#intrate), [`RECEIVED`](doc/functions/financial.md#received), [`PRICEDISC`](doc/functions/financial.md#pricedisc), [`PRICEMAT`](doc/functions/financial.md#pricemat), [`TBILLEQ`](doc/functions/financial.md#tbilleq), [`TBILLPRICE`](doc/functions/financial.md#tbillprice), [`TBILLYIELD`](doc/functions/financial.md#tbillyield), [`DOLLARDE`](doc/functions/financial.md#dollarde), [`DOLLARFR`](doc/functions/financial.md#dollarfr), [`EFFECT`](doc/functions/financial.md#effect), [`NOMINAL`](doc/functions/financial.md#nominal), [`PDURATION`](doc/functions/financial.md#pduration), [`RRI`](doc/functions/financial.md#rri), [`ISPMT`](doc/functions/financial.md#ispmt)

### [Advanced Statistical & Probability (70)](doc/functions/statistical-advanced.md)
[`FISHER`](doc/functions/statistical-advanced.md#fisher), [`FISHERINV`](doc/functions/statistical-advanced.md#fisherinv), [`STANDARDIZE`](doc/functions/statistical-advanced.md#standardize), [`PERMUT`](doc/functions/statistical-advanced.md#permut), [`PERMUTATIONA`](doc/functions/statistical-advanced.md#permutationa), [`DEVSQ`](doc/functions/statistical-advanced.md#devsq), [`KURT`](doc/functions/statistical-advanced.md#kurt), [`SKEW`](doc/functions/statistical-advanced.md#skew), [`SKEW.P`](doc/functions/statistical-advanced.md#skewp), [`COVARIANCE.P`](doc/functions/statistical-advanced.md#covariancep), [`COVARIANCE.S`](doc/functions/statistical-advanced.md#covariances), [`CORREL`](doc/functions/statistical-advanced.md#correl), [`PEARSON`](doc/functions/statistical-advanced.md#pearson), [`RSQ`](doc/functions/statistical-advanced.md#rsq), [`SLOPE`](doc/functions/statistical-advanced.md#slope), [`INTERCEPT`](doc/functions/statistical-advanced.md#intercept), [`STEYX`](doc/functions/statistical-advanced.md#steyx), [`FORECAST.LINEAR`](doc/functions/statistical-advanced.md#forecastlinear), [`PROB`](doc/functions/statistical-advanced.md#prob), [`MODE.MULT`](doc/functions/statistical-advanced.md#modemult), [`STDEVA`](doc/functions/statistical-advanced.md#stdeva), [`STDEVPA`](doc/functions/statistical-advanced.md#stdevpa), [`VARA`](doc/functions/statistical-advanced.md#vara), [`VARPA`](doc/functions/statistical-advanced.md#varpa), [`GAMMA`](doc/functions/statistical-advanced.md#gamma), [`GAMMALN`](doc/functions/statistical-advanced.md#gammaln), [`GAMMALN.PRECISE`](doc/functions/statistical-advanced.md#gammalnprecise), [`GAUSS`](doc/functions/statistical-advanced.md#gauss), [`PHI`](doc/functions/statistical-advanced.md#phi), [`NORM.S.DIST`](doc/functions/statistical-advanced.md#normsdist), [`NORM.S.INV`](doc/functions/statistical-advanced.md#normsinv), [`NORM.DIST`](doc/functions/statistical-advanced.md#normdist), [`NORM.INV`](doc/functions/statistical-advanced.md#norminv), [`BINOM.DIST`](doc/functions/statistical-advanced.md#binomdist), [`BINOM.INV`](doc/functions/statistical-advanced.md#binominv), [`BINOM.DIST.RANGE`](doc/functions/statistical-advanced.md#binomdistrange), [`NEGBINOM.DIST`](doc/functions/statistical-advanced.md#negbinomdist), [`HYPGEOM.DIST`](doc/functions/statistical-advanced.md#hypgeomdist), [`POISSON.DIST`](doc/functions/statistical-advanced.md#poissondist), [`EXPON.DIST`](doc/functions/statistical-advanced.md#expondist), [`GAMMA.DIST`](doc/functions/statistical-advanced.md#gammadist), [`GAMMA.INV`](doc/functions/statistical-advanced.md#gammainv), [`BETA.DIST`](doc/functions/statistical-advanced.md#betadist), [`BETA.INV`](doc/functions/statistical-advanced.md#betainv), [`CHISQ.DIST`](doc/functions/statistical-advanced.md#chisqdist), [`CHISQ.INV`](doc/functions/statistical-advanced.md#chisqinv), [`CHISQ.DIST.RT`](doc/functions/statistical-advanced.md#chisqdistrt), [`CHISQ.INV.RT`](doc/functions/statistical-advanced.md#chisqinvrt), [`T.DIST`](doc/functions/statistical-advanced.md#tdist), [`T.INV`](doc/functions/statistical-advanced.md#tinv), [`T.DIST.2T`](doc/functions/statistical-advanced.md#tdist2t), [`T.INV.2T`](doc/functions/statistical-advanced.md#tinv2t), [`T.DIST.RT`](doc/functions/statistical-advanced.md#tdistrt), [`F.DIST`](doc/functions/statistical-advanced.md#fdist), [`F.INV`](doc/functions/statistical-advanced.md#finv), [`F.DIST.RT`](doc/functions/statistical-advanced.md#fdistrt), [`F.INV.RT`](doc/functions/statistical-advanced.md#finvrt), [`WEIBULL.DIST`](doc/functions/statistical-advanced.md#weibulldist), [`LOGNORM.DIST`](doc/functions/statistical-advanced.md#lognormdist), [`LOGNORM.INV`](doc/functions/statistical-advanced.md#lognorminv), [`CONFIDENCE.NORM`](doc/functions/statistical-advanced.md#confidencenorm), [`CONFIDENCE.T`](doc/functions/statistical-advanced.md#confidencet), [`Z.TEST`](doc/functions/statistical-advanced.md#ztest), [`T.TEST`](doc/functions/statistical-advanced.md#ttest), [`CHISQ.TEST`](doc/functions/statistical-advanced.md#chisqtest), [`F.TEST`](doc/functions/statistical-advanced.md#ftest), [`LINEST`](doc/functions/statistical-advanced.md#linest), [`LOGEST`](doc/functions/statistical-advanced.md#logest), [`TREND`](doc/functions/statistical-advanced.md#trend), [`GROWTH`](doc/functions/statistical-advanced.md#growth)

### [Engineering (54)](doc/functions/engineering.md)
[`DELTA`](doc/functions/engineering.md#delta), [`GESTEP`](doc/functions/engineering.md#gestep), [`BITAND`](doc/functions/engineering.md#bitand), [`BITOR`](doc/functions/engineering.md#bitor), [`BITXOR`](doc/functions/engineering.md#bitxor), [`BITLSHIFT`](doc/functions/engineering.md#bitlshift), [`BITRSHIFT`](doc/functions/engineering.md#bitrshift), [`BIN2DEC`](doc/functions/engineering.md#bin2dec), [`BIN2HEX`](doc/functions/engineering.md#bin2hex), [`BIN2OCT`](doc/functions/engineering.md#bin2oct), [`DEC2BIN`](doc/functions/engineering.md#dec2bin), [`DEC2HEX`](doc/functions/engineering.md#dec2hex), [`DEC2OCT`](doc/functions/engineering.md#dec2oct), [`HEX2BIN`](doc/functions/engineering.md#hex2bin), [`HEX2DEC`](doc/functions/engineering.md#hex2dec), [`HEX2OCT`](doc/functions/engineering.md#hex2oct), [`OCT2BIN`](doc/functions/engineering.md#oct2bin), [`OCT2DEC`](doc/functions/engineering.md#oct2dec), [`OCT2HEX`](doc/functions/engineering.md#oct2hex), [`BASE`](doc/functions/engineering.md#base), [`DECIMAL`](doc/functions/engineering.md#decimal), [`ARABIC`](doc/functions/engineering.md#arabic), [`ROMAN`](doc/functions/engineering.md#roman), [`ERF`](doc/functions/engineering.md#erf), [`ERF.PRECISE`](doc/functions/engineering.md#erfprecise), [`ERFC`](doc/functions/engineering.md#erfc), [`ERFC.PRECISE`](doc/functions/engineering.md#erfcprecise), [`COMPLEX`](doc/functions/engineering.md#complex), [`IMREAL`](doc/functions/engineering.md#imreal), [`IMAGINARY`](doc/functions/engineering.md#imaginary), [`IMABS`](doc/functions/engineering.md#imabs), [`IMARGUMENT`](doc/functions/engineering.md#imargument), [`IMCONJUGATE`](doc/functions/engineering.md#imconjugate), [`IMSUM`](doc/functions/engineering.md#imsum), [`IMSUB`](doc/functions/engineering.md#imsub), [`IMPRODUCT`](doc/functions/engineering.md#improduct), [`IMDIV`](doc/functions/engineering.md#imdiv), [`IMPOWER`](doc/functions/engineering.md#impower), [`IMSQRT`](doc/functions/engineering.md#imsqrt), [`IMEXP`](doc/functions/engineering.md#imexp), [`IMLN`](doc/functions/engineering.md#imln), [`IMLOG10`](doc/functions/engineering.md#imlog10), [`IMLOG2`](doc/functions/engineering.md#imlog2), [`IMSIN`](doc/functions/engineering.md#imsin), [`IMCOS`](doc/functions/engineering.md#imcos), [`IMTAN`](doc/functions/engineering.md#imtan), [`IMSINH`](doc/functions/engineering.md#imsinh), [`IMCOSH`](doc/functions/engineering.md#imcosh), [`IMSEC`](doc/functions/engineering.md#imsec), [`IMSECH`](doc/functions/engineering.md#imsech), [`IMCSC`](doc/functions/engineering.md#imcsc), [`IMCSCH`](doc/functions/engineering.md#imcsch), [`IMCOT`](doc/functions/engineering.md#imcot), [`CONVERT`](doc/functions/engineering.md#convert)

### [Database (12)](doc/functions/database.md)
[`DSUM`](doc/functions/database.md#dsum), [`DAVERAGE`](doc/functions/database.md#daverage), [`DCOUNT`](doc/functions/database.md#dcount), [`DCOUNTA`](doc/functions/database.md#dcounta), [`DMAX`](doc/functions/database.md#dmax), [`DMIN`](doc/functions/database.md#dmin), [`DGET`](doc/functions/database.md#dget), [`DPRODUCT`](doc/functions/database.md#dproduct), [`DSTDEV`](doc/functions/database.md#dstdev), [`DSTDEVP`](doc/functions/database.md#dstdevp), [`DVAR`](doc/functions/database.md#dvar), [`DVARP`](doc/functions/database.md#dvarp)

### [Lambda & Higher-Order (9)](doc/functions/lambda.md)
[`LAMBDA`](doc/functions/lambda.md#lambda), [`LET`](doc/functions/lambda.md#let), [`MAP`](doc/functions/lambda.md#map), [`REDUCE`](doc/functions/lambda.md#reduce), [`SCAN`](doc/functions/lambda.md#scan), [`MAKEARRAY`](doc/functions/lambda.md#makearray), [`BYCOL`](doc/functions/lambda.md#bycol), [`BYROW`](doc/functions/lambda.md#byrow), [`ISOMITTED`](doc/functions/lambda.md#isomitted)

### [Web & Regex (4)](doc/functions/web.md)
[`ENCODEURL`](doc/functions/web.md#encodeurl), [`REGEXMATCH`](doc/functions/web.md#regexmatch), [`REGEXEXTRACT`](doc/functions/web.md#regexextract), [`REGEXREPLACE`](doc/functions/web.md#regexreplace)

## Custom Functions

```dart
class DiscountFunction extends FormulaFunction {
  @override String get name => 'DISCOUNT';
  @override int get minArgs => 2;
  @override int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final price = values[0].toNumber() ?? 0;
    final rate = values[1].toNumber() ?? 0;
    return FormulaValue.number(price * (1 - rate));
  }
}

engine.registerFunction(DiscountFunction());
engine.evaluateString('=DISCOUNT(100, 0.2)', context);
// result = NumberValue(80)
```

## Dependency Tracking

Track which cells depend on which, and determine recalculation order:

```dart
final graph = DependencyGraph();

// B1 = A1 + 1
graph.updateDependencies('B1'.a1, {'A1'.a1});
// C1 = B1 * 2
graph.updateDependencies('C1'.a1, {'B1'.a1});

// When A1 changes, recalculate in order:
final toRecalc = graph.getCellsToRecalculate('A1'.a1);
// [B1, C1]

// Detect circular references
graph.hasCircularReference('A1'.a1); // false
```

## TEXT Format Codes

The `TEXT` function supports Excel-style format codes:

| Format | Example | Result |
|--------|---------|--------|
| `0.00` | `TEXT(3.14159, "0.00")` | `3.14` |
| `#,##0` | `TEXT(1234567, "#,##0")` | `1,234,567` |
| `0%` | `TEXT(0.75, "0%")` | `75%` |
| `000` | `TEXT(5, "000")` | `005` |
| `0.0E+0` | `TEXT(1234, "0.0E+0")` | `1.2E+3` |

## Supported Operators

| Operator | Description | Example |
|----------|-------------|---------|
| `+` `-` `*` `/` | Arithmetic | `=A1+B1*2` |
| `^` | Power | `=2^10` |
| `&` | Concatenation | `="Hello"&" "&"World"` |
| `=` `<>` `<` `>` `<=` `>=` | Comparison | `=A1>10` |
| `-` (prefix) | Negation | `=-A1` |
| `%` (postfix) | Percent | `=50%` |

## Error Types

All Excel-compatible error types are supported: `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A`, `#NULL!`, `#CALC!`, `#CIRCULAR!`

## Examples

### Standalone Dart

[`example/worksheet_formula_example.dart`](example/worksheet_formula_example.dart) -- a pure Dart example demonstrating parsing, evaluation, cell references, dependency tracking, custom functions, and conditional logic. No Flutter required.

```bash
dart run example/worksheet_formula_example.dart
```

### Flutter + Worksheet Widget

[`example/worksheet_integration/`](example/worksheet_integration/) -- a Flutter app integrating `worksheet_formula` with the [`worksheet`](https://pub.dev/packages/worksheet) widget. Shows formula cells evaluated live in a spreadsheet grid with dependency tracking, caching, and a custom `DISCOUNT` function.

```bash
cd example/worksheet_integration
flutter run
```

See the [integration README](example/worksheet_integration/README.md) for architecture details.

## License

See [LICENSE](LICENSE) file.
