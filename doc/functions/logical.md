# Logical Functions

## IF

`IF(logical_test, value_if_true, [value_if_false])`

Returns one value if a condition is TRUE and another if FALSE.

**Example:**
```
=IF(1>0, "yes", "no")    → "yes"
```

## AND

`AND(logical1, [logical2], ...)`

Returns TRUE if all arguments are TRUE.

**Example:**
```
=AND(1>0, 2>1, 3>2)    → TRUE
```

## OR

`OR(logical1, [logical2], ...)`

Returns TRUE if any argument is TRUE.

**Example:**
```
=OR(1>2, 3>2)    → TRUE
```

## NOT

`NOT(logical)`

Reverses the logic of its argument.

**Example:**
```
=NOT(FALSE)    → TRUE
```

## IFERROR

`IFERROR(value, value_if_error)`

Returns value_if_error if value is an error.

**Example:**
```
=IFERROR(1/0, "Error occurred")    → "Error occurred"
```

## IFNA

`IFNA(value, value_if_na)`

Returns value_if_na if value is #N/A.

**Example:**
```
=IFNA(VLOOKUP("x", A1:B5, 2, FALSE), "Not found")    → "Not found"
```

## TRUE

`TRUE()`

Returns the logical value TRUE.

**Example:**
```
=TRUE()    → TRUE
```

## FALSE

`FALSE()`

Returns the logical value FALSE.

**Example:**
```
=FALSE()    → FALSE
```

## IFS

`IFS(condition1, value1, [condition2, value2], ...)`

Chained IF.

**Example:**
```
=IFS(85>=90, "A", 85>=80, "B", 85>=70, "C")    → "B"
```

## SWITCH

`SWITCH(expression, val1, result1, ..., [default])`

Match against cases.

**Example:**
```
=SWITCH(2, 1, "one", 2, "two", 3, "three")    → "two"
```

## XOR

`XOR(logical1, [logical2], ...)`

Exclusive OR.

**Example:**
```
=XOR(TRUE, FALSE, FALSE)    → TRUE
```
