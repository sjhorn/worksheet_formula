# Lambda & Higher-Order Functions

## LAMBDA

`LAMBDA(param1, ..., paramN, body)`

Creates a reusable function.

**Example:**
```
=LAMBDA(x, x*2)(5)    → 10
```

## LET

`LET(name1, value1, ..., nameN, valueN, expr)`

Binds names to values.

**Example:**
```
=LET(x, 10, y, 20, x+y)    → 30
```

## MAP

`MAP(array, lambda)`

Applies lambda to each element of array.

**Example:**
```
=MAP({1,2,3}, LAMBDA(x, x^2))    → {1,4,9}
```

## REDUCE

`REDUCE(initial_value, array, lambda)`

Folds array with lambda(acc, elem).

**Example:**
```
=REDUCE(0, {1,2,3,4}, LAMBDA(a, b, a+b))    → 10
```

## SCAN

`SCAN(initial_value, array, lambda)`

Running fold, returns array of intermediate accumulator values.

**Example:**
```
=SCAN(0, {1,2,3,4}, LAMBDA(a, b, a+b))    → {1,3,6,10}
```

## MAKEARRAY

`MAKEARRAY(rows, cols, lambda)`

Builds array via lambda(row, col).

**Example:**
```
=MAKEARRAY(2, 3, LAMBDA(r, c, r*c))    → {1,2,3;2,4,6}
```

## BYCOL

`BYCOL(array, lambda)`

Applies lambda to each column of the array.

**Example:**
```
=BYCOL({1,2;3,4}, LAMBDA(col, SUM(col)))    → {4,6}
```

## BYROW

`BYROW(array, lambda)`

Applies lambda to each row of the array.

**Example:**
```
=BYROW({1,2;3,4}, LAMBDA(row, SUM(row)))    → {3;7}
```

## ISOMITTED

`ISOMITTED(value)`

Returns TRUE if value is an OmittedValue sentinel.

**Example:**
```
=LAMBDA(x, IF(ISOMITTED(x), "missing", x))()    → "missing"
```
