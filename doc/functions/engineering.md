# Engineering Functions

## DELTA

`DELTA(number1, [number2])`

Tests whether two values are equal. Returns 1 if equal, 0 otherwise.

**Example:**
```
=DELTA(5, 5)    → 1
```

## GESTEP

`GESTEP(number, [step])`

Tests whether a number is greater than or equal to a step value. Returns 1 if number >= step, 0 otherwise.

**Example:**
```
=GESTEP(5, 4)    → 1
```

## BITAND

`BITAND(number1, number2)`

Returns a bitwise AND of two non-negative integers.

**Example:**
```
=BITAND(13, 25)    → 9
```

## BITOR

`BITOR(number1, number2)`

Returns a bitwise OR of two non-negative integers.

**Example:**
```
=BITOR(13, 25)    → 29
```

## BITXOR

`BITXOR(number1, number2)`

Returns a bitwise XOR of two non-negative integers.

**Example:**
```
=BITXOR(13, 25)    → 20
```

## BITLSHIFT

`BITLSHIFT(number, shift_amount)`

Returns a number shifted left by the specified number of bits.

**Example:**
```
=BITLSHIFT(4, 2)    → 16
```

## BITRSHIFT

`BITRSHIFT(number, shift_amount)`

Returns a number shifted right by the specified number of bits.

**Example:**
```
=BITRSHIFT(16, 2)    → 4
```

## BIN2DEC

`BIN2DEC(number)`

Converts a binary number to decimal.

**Example:**
```
=BIN2DEC("1100100")    → 100
```

## BIN2HEX

`BIN2HEX(number, [places])`

Converts a binary number to hexadecimal.

**Example:**
```
=BIN2HEX("11111011", 4)    → "00FB"
```

## BIN2OCT

`BIN2OCT(number, [places])`

Converts a binary number to octal.

**Example:**
```
=BIN2OCT("1001", 4)    → "0011"
```

## DEC2BIN

`DEC2BIN(number, [places])`

Converts a decimal number to binary.

**Example:**
```
=DEC2BIN(9, 4)    → "1001"
```

## DEC2HEX

`DEC2HEX(number, [places])`

Converts a decimal number to hexadecimal.

**Example:**
```
=DEC2HEX(255, 4)    → "00FF"
```

## DEC2OCT

`DEC2OCT(number, [places])`

Converts a decimal number to octal.

**Example:**
```
=DEC2OCT(58, 4)    → "0072"
```

## HEX2BIN

`HEX2BIN(number, [places])`

Converts a hexadecimal number to binary.

**Example:**
```
=HEX2BIN("F", 8)    → "00001111"
```

## HEX2DEC

`HEX2DEC(number)`

Converts a hexadecimal number to decimal.

**Example:**
```
=HEX2DEC("FF")    → 255
```

## HEX2OCT

`HEX2OCT(number, [places])`

Converts a hexadecimal number to octal.

**Example:**
```
=HEX2OCT("F", 4)    → "0017"
```

## OCT2BIN

`OCT2BIN(number, [places])`

Converts an octal number to binary.

**Example:**
```
=OCT2BIN("7", 4)    → "0111"
```

## OCT2DEC

`OCT2DEC(number)`

Converts an octal number to decimal.

**Example:**
```
=OCT2DEC("77")    → 63
```

## OCT2HEX

`OCT2HEX(number, [places])`

Converts an octal number to hexadecimal.

**Example:**
```
=OCT2HEX("77", 4)    → "003F"
```

## BASE

`BASE(number, radix, [min_length])`

Converts a number into a text representation with the given radix (base).

**Example:**
```
=BASE(255, 16)    → "FF"
```

## DECIMAL

`DECIMAL(text, radix)`

Converts a text representation of a number in a given base into a decimal number.

**Example:**
```
=DECIMAL("FF", 16)    → 255
```

## ARABIC

`ARABIC(text)`

Converts a Roman numeral to an Arabic numeral.

**Example:**
```
=ARABIC("MCMXCIV")    → 1994
```

## ROMAN

`ROMAN(number, [form])`

Converts an Arabic numeral to a Roman numeral as text.

**Example:**
```
=ROMAN(1994)    → "MCMXCIV"
```

## ERF

`ERF(lower_limit, [upper_limit])`

Returns the error function integrated between the given limits.

**Example:**
```
=ERF(1)    → 0.842700793
```

## ERF.PRECISE

`ERF.PRECISE(x)`

Returns the error function integrated between 0 and the given limit.

**Example:**
```
=ERF.PRECISE(1)    → 0.842700793
```

## ERFC

`ERFC(x)`

Returns the complementary error function integrated between x and infinity.

**Example:**
```
=ERFC(1)    → 0.157299207
```

## ERFC.PRECISE

`ERFC.PRECISE(x)`

Returns the complementary error function integrated between x and infinity.

**Example:**
```
=ERFC.PRECISE(1)    → 0.157299207
```

## COMPLEX

`COMPLEX(real_num, i_num, [suffix])`

Converts real and imaginary coefficients into a complex number of the form a+bi or a+bj.

**Example:**
```
=COMPLEX(3, 4)    → "3+4i"
```

## IMREAL

`IMREAL(inumber)`

Returns the real coefficient of a complex number in x+yi text format.

**Example:**
```
=IMREAL("3+4i")    → 3
```

## IMAGINARY

`IMAGINARY(inumber)`

Returns the imaginary coefficient of a complex number in x+yi text format.

**Example:**
```
=IMAGINARY("3+4i")    → 4
```

## IMABS

`IMABS(inumber)`

Returns the absolute value (modulus) of a complex number.

**Example:**
```
=IMABS("3+4i")    → 5
```

## IMARGUMENT

`IMARGUMENT(inumber)`

Returns the argument (theta) of a complex number, an angle expressed in radians.

**Example:**
```
=IMARGUMENT("3+4i")    → 0.927295218
```

## IMCONJUGATE

`IMCONJUGATE(inumber)`

Returns the complex conjugate of a complex number.

**Example:**
```
=IMCONJUGATE("3+4i")    → "3-4i"
```

## IMSUM

`IMSUM(inumber1, [inumber2], ...)`

Returns the sum of two or more complex numbers.

**Example:**
```
=IMSUM("3+4i", "1+2i")    → "4+6i"
```

## IMSUB

`IMSUB(inumber1, inumber2)`

Returns the difference of two complex numbers.

**Example:**
```
=IMSUB("13+4i", "5+3i")    → "8+i"
```

## IMPRODUCT

`IMPRODUCT(inumber1, [inumber2], ...)`

Returns the product of two or more complex numbers.

**Example:**
```
=IMPRODUCT("3+4i", "5-3i")    → "27+11i"
```

## IMDIV

`IMDIV(inumber1, inumber2)`

Returns the quotient of two complex numbers.

**Example:**
```
=IMDIV("2+4i", "1+i")    → "3+i"
```

## IMPOWER

`IMPOWER(inumber, number)`

Returns a complex number raised to an integer power.

**Example:**
```
=IMPOWER("2+3i", 2)    → "-5+12i"
```

## IMSQRT

`IMSQRT(inumber)`

Returns the square root of a complex number.

**Example:**
```
=IMSQRT("4")    → "2"
```

## IMEXP

`IMEXP(inumber)`

Returns the exponential of a complex number.

**Example:**
```
=IMEXP("1+i")    → "1.46869393991589+2.28735528717884i"
```

## IMLN

`IMLN(inumber)`

Returns the natural logarithm of a complex number.

**Example:**
```
=IMLN("3+4i")    → "1.6094379124341+0.927295218001612i"
```

## IMLOG10

`IMLOG10(inumber)`

Returns the base-10 logarithm of a complex number.

**Example:**
```
=IMLOG10("3+4i")    → "0.698970004336019+0.402719196273373i"
```

## IMLOG2

`IMLOG2(inumber)`

Returns the base-2 logarithm of a complex number.

**Example:**
```
=IMLOG2("3+4i")    → "2.32192809488736+1.33780421245098i"
```

## IMSIN

`IMSIN(inumber)`

Returns the sine of a complex number.

**Example:**
```
=IMSIN("3+4i")    → "3.85373803791938-27.0168132580039i"
```

## IMCOS

`IMCOS(inumber)`

Returns the cosine of a complex number.

**Example:**
```
=IMCOS("1+i")    → "0.83373002513115-0.988897705762865i"
```

## IMTAN

`IMTAN(inumber)`

Returns the tangent of a complex number.

**Example:**
```
=IMTAN("1+i")    → "0.271752585319512+1.08392332733869i"
```

## IMSINH

`IMSINH(inumber)`

Returns the hyperbolic sine of a complex number.

**Example:**
```
=IMSINH("1+i")    → "0.634963914784736+1.29845758141598i"
```

## IMCOSH

`IMCOSH(inumber)`

Returns the hyperbolic cosine of a complex number.

**Example:**
```
=IMCOSH("1+i")    → "0.83373002513115+0.988897705762865i"
```

## IMSEC

`IMSEC(inumber)`

Returns the secant of a complex number.

**Example:**
```
=IMSEC("1+i")    → "0.498337030555187+0.591083841721045i"
```

## IMSECH

`IMSECH(inumber)`

Returns the hyperbolic secant of a complex number.

**Example:**
```
=IMSECH("1+i")    → "0.498337030555187-0.591083841721045i"
```

## IMCSC

`IMCSC(inumber)`

Returns the cosecant of a complex number.

**Example:**
```
=IMCSC("1+i")    → "0.621518017170428-0.303931001628426i"
```

## IMCSCH

`IMCSCH(inumber)`

Returns the hyperbolic cosecant of a complex number.

**Example:**
```
=IMCSCH("1+i")    → "0.303931001628426-0.621518017170428i"
```

## IMCOT

`IMCOT(inumber)`

Returns the cotangent of a complex number.

**Example:**
```
=IMCOT("1+i")    → "0.217621561854403-0.868014142895925i"
```

## CONVERT

`CONVERT(number, from_unit, to_unit)`

Converts a number from one measurement system to another.

**Example:**
```
=CONVERT(1, "mi", "km")    → 1.609344
```
