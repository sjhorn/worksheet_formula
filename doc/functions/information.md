# Information Functions

## ISBLANK

`ISBLANK(value)`

Returns TRUE if value is empty.

**Example:**
```
=ISBLANK("")    → TRUE
```

## ISERROR

`ISERROR(value)`

Returns TRUE if value is any error.

**Example:**
```
=ISERROR(1/0)    → TRUE
```

## ISNUMBER

`ISNUMBER(value)`

Returns TRUE if value is a number.

**Example:**
```
=ISNUMBER(42)    → TRUE
```

## ISTEXT

`ISTEXT(value)`

Returns TRUE if value is text.

**Example:**
```
=ISTEXT("hello")    → TRUE
```

## ISLOGICAL

`ISLOGICAL(value)`

Returns TRUE if value is a boolean.

**Example:**
```
=ISLOGICAL(TRUE)    → TRUE
```

## ISNA

`ISNA(value)`

Returns TRUE if value is #N/A.

**Example:**
```
=ISNA(NA())    → TRUE
```

## TYPE

`TYPE(value)`

Returns the type of a value.

**Example:**
```
=TYPE("hello")    → 2
```

## ISERR

`ISERR(value)`

Returns TRUE if value is an error other than #N/A.

**Example:**
```
=ISERR(1/0)    → TRUE
```

## ISNONTEXT

`ISNONTEXT(value)`

Returns TRUE if value is not text.

**Example:**
```
=ISNONTEXT(123)    → TRUE
```

## ISEVEN

`ISEVEN(number)`

Returns TRUE if number is even.

**Example:**
```
=ISEVEN(4)    → TRUE
```

## ISODD

`ISODD(number)`

Returns TRUE if number is odd.

**Example:**
```
=ISODD(3)    → TRUE
```

## ISREF

`ISREF(value)`

Returns TRUE if value is a cell or range reference.

**Example:**
```
=ISREF(A1)    → TRUE
```

## N

`N(value)`

Converts a value to a number.

**Example:**
```
=N(TRUE)    → 1
```

## NA

`NA()`

Returns the #N/A error value.

**Example:**
```
=NA()    → #N/A
```

## ERROR.TYPE

`ERROR.TYPE(error_val)`

Returns a number corresponding to an error type.

**Example:**
```
=ERROR.TYPE(#REF!)    → 4
```
