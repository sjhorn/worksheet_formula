# Financial Functions

## EFFECT

`EFFECT(nominal_rate, npery)`

Effective annual interest rate.

**Example:**
```
=EFFECT(0.06, 4)    → 0.06136
```

## NOMINAL

`NOMINAL(effect_rate, npery)`

Nominal annual interest rate.

**Example:**
```
=NOMINAL(0.06136, 4)    → 0.06
```

## PDURATION

`PDURATION(rate, pv, fv)`

Periods required for investment to reach a value.

**Example:**
```
=PDURATION(0.05, 1000, 2000)    → 14.21
```

## RRI

`RRI(nper, pv, fv)`

Equivalent interest rate for growth of investment.

**Example:**
```
=RRI(10, 1000, 2000)    → 0.07177
```

## ISPMT

`ISPMT(rate, per, nper, pv)`

Interest paid for a given period.

**Example:**
```
=ISPMT(0.05/12, 1, 36, 8000000)    → -27777.78
```

## SLN

`SLN(cost, salvage, life)`

Straight-line depreciation.

**Example:**
```
=SLN(30000, 7500, 10)    → 2250
```

## SYD

`SYD(cost, salvage, life, per)`

Sum-of-years-digits depreciation.

**Example:**
```
=SYD(30000, 7500, 10, 1)    → 4090.91
```

## DOLLARDE

`DOLLARDE(fractional_dollar, fraction)`

Convert fractional to decimal.

**Example:**
```
=DOLLARDE(1.02, 16)    → 1.125
```

## DOLLARFR

`DOLLARFR(decimal_dollar, fraction)`

Convert decimal to fractional.

**Example:**
```
=DOLLARFR(1.125, 16)    → 1.02
```

## FVSCHEDULE

`FVSCHEDULE(principal, schedule)`

Future value with variable rates.

**Example:**
```
=FVSCHEDULE(10000, {0.05, 0.06, 0.07})    → 11892.46
```

## NPV

`NPV(rate, value1, [value2], ...)`

Net present value.

**Example:**
```
=NPV(0.08, -40000, 8000, 9200, 10000, 12000, 14500)    → 1922.06
```

## TBILLEQ

`TBILLEQ(settlement, maturity, discount)`

T-Bill bond-equivalent yield.

**Example:**
```
=TBILLEQ(40000, 40182, 0.09)    → 0.09415
```

## TBILLPRICE

`TBILLPRICE(settlement, maturity, discount)`

T-Bill price per $100.

**Example:**
```
=TBILLPRICE(40000, 40182, 0.09)    → 95.45
```

## TBILLYIELD

`TBILLYIELD(settlement, maturity, pr)`

T-Bill yield.

**Example:**
```
=TBILLYIELD(40000, 40182, 98.45)    → 0.03138
```

## PMT

`PMT(rate, nper, pv, [fv], [type])`

Payment for a loan.

**Example:**
```
=PMT(0.05/12, 360, -200000)    → 1073.64
```

## FV

`FV(rate, nper, pmt, [pv], [type])`

Future value.

**Example:**
```
=FV(0.06/12, 120, -200)    → 32776.87
```

## PV

`PV(rate, nper, pmt, [fv], [type])`

Present value.

**Example:**
```
=PV(0.08/12, 240, -500)    → 59777.15
```

## NPER

`NPER(rate, pmt, pv, [fv], [type])`

Number of periods.

**Example:**
```
=NPER(0.06/12, -100, 0, 10000)    → 81.29
```

## IPMT

`IPMT(rate, per, nper, pv, [fv], [type])`

Interest payment for a period.

**Example:**
```
=IPMT(0.05/12, 1, 360, 200000)    → -833.33
```

## PPMT

`PPMT(rate, per, nper, pv, [fv], [type])`

Principal payment for a period.

**Example:**
```
=PPMT(0.05/12, 1, 360, 200000)    → -240.31
```

## CUMIPMT

`CUMIPMT(rate, nper, pv, start, end, type)`

Cumulative interest.

**Example:**
```
=CUMIPMT(0.05/12, 360, 200000, 1, 12, 0)    → -9916.77
```

## CUMPRINC

`CUMPRINC(rate, nper, pv, start, end, type)`

Cumulative principal.

**Example:**
```
=CUMPRINC(0.05/12, 360, 200000, 1, 12, 0)    → -2966.88
```

## RATE

`RATE(nper, pmt, pv, [fv], [type], [guess])`

Interest rate per period.

**Example:**
```
=RATE(360, -1073.64, 200000)    → 0.00417
```

## IRR

`IRR(values, [guess])`

Internal rate of return.

**Example:**
```
=IRR({-70000, 12000, 15000, 18000, 21000, 26000})    → 0.08663
```

## XNPV

`XNPV(rate, values, dates)`

NPV for irregular cash flows.

**Example:**
```
=XNPV(0.09, {-10000, 2750, 4250, 3250, 2750}, {39448, 39508, 39691, 39873, 40057})    → 2086.65
```

## XIRR

`XIRR(values, dates, [guess])`

IRR for irregular cash flows.

**Example:**
```
=XIRR({-10000, 2750, 4250, 3250, 2750}, {39448, 39508, 39691, 39873, 40057})    → 0.37336
```

## MIRR

`MIRR(values, finance_rate, reinvest_rate)`

Modified internal rate of return.

**Example:**
```
=MIRR({-120000, 39000, 30000, 21000, 37000, 46000}, 0.10, 0.12)    → 0.12609
```

## DB

`DB(cost, salvage, life, period, [month])`

Fixed declining balance.

**Example:**
```
=DB(1000000, 100000, 6, 1, 7)    → 186083.33
```

## DDB

`DDB(cost, salvage, life, period, [factor])`

Double declining balance.

**Example:**
```
=DDB(10000, 1000, 5, 1)    → 4000
```

## VDB

`VDB(cost, salvage, life, start_period, end_period, [factor], [no_switch])`

Variable declining balance depreciation between any two periods.

**Example:**
```
=VDB(10000, 1000, 5, 0, 1)    → 4000
```

## DISC

`DISC(settlement, maturity, pr, redemption, [basis])`

Discount rate for a security.

**Example:**
```
=DISC(40000, 40182, 97.975, 100)    → 0.04013
```

## INTRATE

`INTRATE(settlement, maturity, investment, redemption, [basis])`

Interest rate for a fully invested security.

**Example:**
```
=INTRATE(40000, 40182, 1000000, 1014420)    → 0.02854
```

## RECEIVED

`RECEIVED(settlement, maturity, investment, discount, [basis])`

Amount received at maturity for a fully invested security.

**Example:**
```
=RECEIVED(40000, 40182, 1000000, 0.05)    → 1025316.46
```

## PRICEDISC

`PRICEDISC(settlement, maturity, discount, redemption, [basis])`

Price per $100 face value of a discounted security.

**Example:**
```
=PRICEDISC(40000, 40182, 0.05, 100)    → 97.47
```

## PRICEMAT

`PRICEMAT(settlement, maturity, issue, rate, yld, [basis])`

Price per $100 face value of a security that pays interest at maturity.

**Example:**
```
=PRICEMAT(40000, 40182, 39814, 0.06, 0.05)    → 100.50
```

## ACCRINT

`ACCRINT(issue, first_interest, settlement, rate, par, frequency, [basis])`

Accrued interest for a security that pays periodic interest.

**Example:**
```
=ACCRINT(39814, 39994, 40000, 0.06, 1000, 2)    → 30.58
```

## PRICE

`PRICE(settlement, maturity, rate, yld, redemption, frequency, [basis])`

Price per $100 face value of a security that pays periodic interest.

**Example:**
```
=PRICE(40000, 43831, 0.0575, 0.065, 100, 2)    → 95.04
```

## YIELD

`YIELD(settlement, maturity, rate, pr, redemption, frequency, [basis])`

Yield on a security that pays periodic interest.

**Example:**
```
=YIELD(40000, 43831, 0.0575, 95.04, 100, 2)    → 0.065
```

## DURATION

`DURATION(settlement, maturity, coupon, yld, frequency, [basis])`

Macaulay duration of a security.

**Example:**
```
=DURATION(40000, 43831, 0.08, 0.09, 2)    → 5.99
```

## MDURATION

`MDURATION(settlement, maturity, coupon, yld, frequency, [basis])`

Modified Macaulay duration of a security.

**Example:**
```
=MDURATION(40000, 43831, 0.08, 0.09, 2)    → 5.74
```
