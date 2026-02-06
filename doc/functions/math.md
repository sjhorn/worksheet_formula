# Math & Trigonometry Functions

## SUM

`SUM(number1, [number2], ...)`

Adds all numbers.

**Example:**
```
=SUM(1, 2, 3)    → 6
```

## AVERAGE

`AVERAGE(number1, [number2], ...)`

Returns the average.

**Example:**
```
=AVERAGE(10, 20, 30)    → 20
```

## MIN

`MIN(number1, [number2], ...)`

Returns the minimum value.

**Example:**
```
=MIN(5, 3, 8, 1)    → 1
```

## MAX

`MAX(number1, [number2], ...)`

Returns the maximum value.

**Example:**
```
=MAX(5, 3, 8, 1)    → 8
```

## ABS

`ABS(number)`

Returns the absolute value.

**Example:**
```
=ABS(-7)    → 7
```

## ROUND

`ROUND(number, num_digits)`

Rounds a number.

**Example:**
```
=ROUND(3.14159, 2)    → 3.14
```

## INT

`INT(number)`

Rounds down to the nearest integer.

**Example:**
```
=INT(5.8)    → 5
```

## MOD

`MOD(number, divisor)`

Returns the remainder.

**Example:**
```
=MOD(10, 3)    → 1
```

## SQRT

`SQRT(number)`

Returns the square root.

**Example:**
```
=SQRT(16)    → 4
```

## POWER

`POWER(number, power)`

Returns number raised to a power.

**Example:**
```
=POWER(2, 10)    → 1024
```

## SUMPRODUCT

`SUMPRODUCT(array1, [array2], ...)`

Sum of element-wise products.

**Example:**
```
=SUMPRODUCT({1,2,3}, {4,5,6})    → 32
```

## ROUNDUP

`ROUNDUP(number, num_digits)`

Rounds away from zero.

**Example:**
```
=ROUNDUP(3.141, 2)    → 3.15
```

## ROUNDDOWN

`ROUNDDOWN(number, num_digits)`

Rounds toward zero.

**Example:**
```
=ROUNDDOWN(3.149, 2)    → 3.14
```

## CEILING

`CEILING(number, significance)`

Rounds up to nearest multiple.

**Example:**
```
=CEILING(4.2, 1)    → 5
```

## FLOOR

`FLOOR(number, significance)`

Rounds down to nearest multiple.

**Example:**
```
=FLOOR(4.8, 1)    → 4
```

## SIGN

`SIGN(number)`

Returns -1, 0, or 1.

**Example:**
```
=SIGN(-42)    → -1
```

## PRODUCT

`PRODUCT(number1, [number2], ...)`

Multiplies all numbers.

**Example:**
```
=PRODUCT(2, 3, 4)    → 24
```

## RAND

`RAND()`

Returns random number between 0 and 1.

**Example:**
```
=RAND()    → 0.5281 (varies)
```

## RANDBETWEEN

`RANDBETWEEN(bottom, top)`

Returns random integer in range.

**Example:**
```
=RANDBETWEEN(1, 100)    → 47 (varies)
```

## PI

`PI()`

Returns the mathematical constant pi.

**Example:**
```
=PI()    → 3.14159265358979
```

## LN

`LN(number)`

Returns the natural logarithm.

**Example:**
```
=LN(2.71828)    → 1
```

## LOG

`LOG(number, [base])`

Returns the logarithm with specified base.

**Example:**
```
=LOG(1000, 10)    → 3
```

## LOG10

`LOG10(number)`

Returns the base-10 logarithm.

**Example:**
```
=LOG10(100)    → 2
```

## EXP

`EXP(number)`

Returns e raised to a power.

**Example:**
```
=EXP(1)    → 2.71828182845905
```

## SIN

`SIN(number)`

Returns the sine (radians).

**Example:**
```
=SIN(PI()/2)    → 1
```

## COS

`COS(number)`

Returns the cosine (radians).

**Example:**
```
=COS(0)    → 1
```

## TAN

`TAN(number)`

Returns the tangent (radians).

**Example:**
```
=TAN(PI()/4)    → 1
```

## ASIN

`ASIN(number)`

Returns the arcsine.

**Example:**
```
=ASIN(0.5)    → 0.523598775598299
```

## ACOS

`ACOS(number)`

Returns the arccosine.

**Example:**
```
=ACOS(0.5)    → 1.0471975511966
```

## ATAN

`ATAN(number)`

Returns the arctangent.

**Example:**
```
=ATAN(1)    → 0.785398163397448
```

## ATAN2

`ATAN2(x_num, y_num)`

Returns the arctangent from x and y coordinates.

**Example:**
```
=ATAN2(1, 1)    → 0.785398163397448
```

## DEGREES

`DEGREES(angle)`

Converts radians to degrees.

**Example:**
```
=DEGREES(PI())    → 180
```

## RADIANS

`RADIANS(angle)`

Converts degrees to radians.

**Example:**
```
=RADIANS(180)    → 3.14159265358979
```

## EVEN

`EVEN(number)`

Rounds away from zero to the nearest even integer.

**Example:**
```
=EVEN(3)    → 4
```

## ODD

`ODD(number)`

Rounds away from zero to the nearest odd integer.

**Example:**
```
=ODD(2)    → 3
```

## GCD

`GCD(number1, [number2], ...)`

Returns the greatest common divisor.

**Example:**
```
=GCD(12, 18)    → 6
```

## LCM

`LCM(number1, [number2], ...)`

Returns the least common multiple.

**Example:**
```
=LCM(4, 6)    → 12
```

## TRUNC

`TRUNC(number, [num_digits])`

Truncates toward zero.

**Example:**
```
=TRUNC(3.75, 1)    → 3.7
```

## MROUND

`MROUND(number, multiple)`

Rounds to nearest multiple.

**Example:**
```
=MROUND(13, 5)    → 15
```

## QUOTIENT

`QUOTIENT(numerator, denominator)`

Returns the integer portion of a division.

**Example:**
```
=QUOTIENT(7, 2)    → 3
```

## COMBIN

`COMBIN(n, k)`

Returns the number of combinations.

**Example:**
```
=COMBIN(10, 3)    → 120
```

## COMBINA

`COMBINA(n, k)`

Returns the number of combinations with repetitions.

**Example:**
```
=COMBINA(4, 2)    → 10
```

## FACT

`FACT(number)`

Returns the factorial.

**Example:**
```
=FACT(5)    → 120
```

## FACTDOUBLE

`FACTDOUBLE(number)`

Returns the double factorial.

**Example:**
```
=FACTDOUBLE(7)    → 105
```

## SUMSQ

`SUMSQ(number1, [number2], ...)`

Returns the sum of squares.

**Example:**
```
=SUMSQ(3, 4)    → 25
```

## SUBTOTAL

`SUBTOTAL(function_num, ref1, [ref2], ...)`

Applies a function to data.

**Example:**
```
=SUBTOTAL(9, A1:A10)    → 55
```

## AGGREGATE

`AGGREGATE(function_num, options, ref1, [ref2], ...)`

Applies a function with options to ignore errors.

**Example:**
```
=AGGREGATE(9, 6, A1:A10)    → 55
```

## SERIESSUM

`SERIESSUM(x, n, m, coefficients)`

Power series: sum of a_i * x^(n + i*m).

**Example:**
```
=SERIESSUM(2, 0, 1, {1,1,1})    → 7
```

## SQRTPI

`SQRTPI(number)`

Returns the square root of (number * pi).

**Example:**
```
=SQRTPI(1)    → 1.7724538509055
```

## MULTINOMIAL

`MULTINOMIAL(n1, n2, ...)`

Returns (n1+n2+...)! / (n1!*n2!*...).

**Example:**
```
=MULTINOMIAL(2, 3)    → 10
```
