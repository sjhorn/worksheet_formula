# Advanced Statistical & Probability Functions

## FISHER

`FISHER(x)`

Returns the Fisher transformation.

**Example:**
```
=FISHER(0.75)    → 0.9730
```

## FISHERINV

`FISHERINV(y)`

Returns the inverse Fisher transformation.

**Example:**
```
=FISHERINV(0.9730)    → 0.75
```

## STANDARDIZE

`STANDARDIZE(x, mean, standard_dev)`

Returns the normalized value.

**Example:**
```
=STANDARDIZE(42, 40, 1.5)    → 1.3333
```

## PERMUT

`PERMUT(number, number_chosen)`

Returns the number of permutations.

**Example:**
```
=PERMUT(10, 3)    → 720
```

## PERMUTATIONA

`PERMUTATIONA(number, number_chosen)`

Returns the number of permutations with repetition.

**Example:**
```
=PERMUTATIONA(4, 3)    → 64
```

## DEVSQ

`DEVSQ(number1, [number2], ...)`

Returns the sum of squared deviations.

**Example:**
```
=DEVSQ(4, 5, 8, 7, 11, 4, 3)    → 48
```

## KURT

`KURT(number1, [number2], ...)`

Returns the excess kurtosis.

**Example:**
```
=KURT(3, 4, 5, 2, 3, 4, 5, 6, 4, 7)    → -0.1518
```

## SKEW

`SKEW(number1, [number2], ...)`

Returns the sample skewness.

**Example:**
```
=SKEW(3, 4, 5, 2, 3, 4, 5, 6, 4, 7)    → 0.3595
```

## SKEW.P

`SKEW.P(number1, [number2], ...)`

Returns the population skewness.

**Example:**
```
=SKEW.P(3, 4, 5, 2, 3, 4, 5, 6, 4, 7)    → 0.3033
```

## COVARIANCE.P

`COVARIANCE.P(array1, array2)`

Returns the population covariance.

**Example:**
```
=COVARIANCE.P({3,2,4,5,6}, {9,7,12,15,17})    → 5.2
```

## COVARIANCE.S

`COVARIANCE.S(array1, array2)`

Returns the sample covariance.

**Example:**
```
=COVARIANCE.S({3,2,4,5,6}, {9,7,12,15,17})    → 6.5
```

## CORREL

`CORREL(array1, array2)`

Returns the Pearson correlation coefficient.

**Example:**
```
=CORREL({3,2,4,5,6}, {9,7,12,15,17})    → 0.9971
```

## PEARSON

`PEARSON(array1, array2)`

Same as CORREL.

**Example:**
```
=PEARSON({3,2,4,5,6}, {9,7,12,15,17})    → 0.9971
```

## RSQ

`RSQ(known_ys, known_xs)`

Returns the R-squared value.

**Example:**
```
=RSQ({2,3,9,1,8}, {6,5,11,7,5})    → 0.4060
```

## SLOPE

`SLOPE(known_ys, known_xs)`

Returns the slope of the linear regression line.

**Example:**
```
=SLOPE({2,3,9,1,8}, {6,5,11,7,5})    → 0.6842
```

## INTERCEPT

`INTERCEPT(known_ys, known_xs)`

Returns the intercept of the linear regression line.

**Example:**
```
=INTERCEPT({2,3,9,1,8}, {6,5,11,7,5})    → 0.0526
```

## STEYX

`STEYX(known_ys, known_xs)`

Returns the standard error of the regression.

**Example:**
```
=STEYX({2,3,9,1,8}, {6,5,11,7,5})    → 3.3056
```

## FORECAST.LINEAR

`FORECAST.LINEAR(x, known_ys, known_xs)`

Returns the predicted value using linear regression.

**Example:**
```
=FORECAST.LINEAR(30, {6,7,9,15,21}, {20,28,31,38,40})    → 10.6073
```

## PROB

`PROB(x_range, prob_range, lower_limit, [upper_limit])`

Returns the probability that values fall within a range.

**Example:**
```
=PROB({1,2,3,4}, {0.1,0.2,0.3,0.4}, 2, 3)    → 0.5
```

## MODE.MULT

`MODE.MULT(number1, [number2], ...)`

Returns all modal values (most frequently occurring).

**Example:**
```
=MODE.MULT(1,2,3,4,3,2,1,2,3)    → {2,3}
```

## STDEVA

`STDEVA(value1, [value2], ...)`

Returns the sample standard deviation including text and logical values.

**Example:**
```
=STDEVA(1, 3, 5, 7, TRUE)    → 2.6077
```

## STDEVPA

`STDEVPA(value1, [value2], ...)`

Returns the population standard deviation including text and logical values.

**Example:**
```
=STDEVPA(1, 3, 5, 7, TRUE)    → 2.3324
```

## VARA

`VARA(value1, [value2], ...)`

Returns the sample variance including text and logical values.

**Example:**
```
=VARA(1, 3, 5, 7, TRUE)    → 6.8
```

## VARPA

`VARPA(value1, [value2], ...)`

Returns the population variance including text and logical values.

**Example:**
```
=VARPA(1, 3, 5, 7, TRUE)    → 5.44
```

## GAMMA

`GAMMA(x)`

Returns the Gamma function value.

**Example:**
```
=GAMMA(4)    → 6
```

## GAMMALN

`GAMMALN(x)`

Returns the natural log of the Gamma function.

**Example:**
```
=GAMMALN(4)    → 1.7918
```

## GAMMALN.PRECISE

`GAMMALN.PRECISE(x)`

Same as GAMMALN.

**Example:**
```
=GAMMALN.PRECISE(4)    → 1.7918
```

## GAUSS

`GAUSS(z)`

Returns the probability between the mean and z standard deviations.

**Example:**
```
=GAUSS(2)    → 0.4772
```

## PHI

`PHI(x)`

Returns the value of the standard normal density function.

**Example:**
```
=PHI(0)    → 0.3989
```

## NORM.S.DIST

`NORM.S.DIST(z, cumulative)`

Returns the standard normal distribution.

**Example:**
```
=NORM.S.DIST(1.333333, TRUE)    → 0.9088
```

## NORM.S.INV

`NORM.S.INV(probability)`

Returns the inverse of the standard normal cumulative distribution.

**Example:**
```
=NORM.S.INV(0.9088)    → 1.3333
```

## NORM.DIST

`NORM.DIST(x, mean, standard_dev, cumulative)`

Returns the normal distribution.

**Example:**
```
=NORM.DIST(42, 40, 1.5, TRUE)    → 0.9088
```

## NORM.INV

`NORM.INV(probability, mean, standard_dev)`

Returns the inverse of the normal cumulative distribution.

**Example:**
```
=NORM.INV(0.9088, 40, 1.5)    → 42
```

## BINOM.DIST

`BINOM.DIST(number_s, trials, probability_s, cumulative)`

Returns the binomial distribution probability.

**Example:**
```
=BINOM.DIST(6, 10, 0.5, FALSE)    → 0.2051
```

## BINOM.INV

`BINOM.INV(trials, probability_s, alpha)`

Returns the smallest value for which the cumulative binomial distribution is greater than or equal to alpha.

**Example:**
```
=BINOM.INV(6, 0.5, 0.75)    → 4
```

## BINOM.DIST.RANGE

`BINOM.DIST.RANGE(trials, probability_s, number_s, [number_s2])`

Returns the probability of a trial result using a binomial distribution.

**Example:**
```
=BINOM.DIST.RANGE(60, 0.75, 45, 50)    → 0.5765
```

## NEGBINOM.DIST

`NEGBINOM.DIST(number_f, number_s, probability_s, cumulative)`

Returns the negative binomial distribution.

**Example:**
```
=NEGBINOM.DIST(10, 5, 0.25, FALSE)    → 0.0550
```

## HYPGEOM.DIST

`HYPGEOM.DIST(sample_s, number_sample, population_s, number_pop, cumulative)`

Returns the hypergeometric distribution.

**Example:**
```
=HYPGEOM.DIST(1, 4, 8, 20, FALSE)    → 0.3633
```

## POISSON.DIST

`POISSON.DIST(x, mean, cumulative)`

Returns the Poisson distribution.

**Example:**
```
=POISSON.DIST(2, 5, FALSE)    → 0.0842
```

## EXPON.DIST

`EXPON.DIST(x, lambda, cumulative)`

Returns the exponential distribution.

**Example:**
```
=EXPON.DIST(0.2, 10, TRUE)    → 0.8647
```

## GAMMA.DIST

`GAMMA.DIST(x, alpha, beta, cumulative)`

Returns the gamma distribution.

**Example:**
```
=GAMMA.DIST(10.00001131, 9, 2, TRUE)    → 0.0680
```

## GAMMA.INV

`GAMMA.INV(probability, alpha, beta)`

Returns the inverse of the gamma cumulative distribution.

**Example:**
```
=GAMMA.INV(0.068, 9, 2)    → 10.0000
```

## BETA.DIST

`BETA.DIST(x, alpha, beta, cumulative, [A], [B])`

Returns the beta distribution.

**Example:**
```
=BETA.DIST(2, 8, 10, TRUE, 1, 3)    → 0.6854
```

## BETA.INV

`BETA.INV(probability, alpha, beta, [A], [B])`

Returns the inverse of the beta cumulative distribution.

**Example:**
```
=BETA.INV(0.6854, 8, 10, 1, 3)    → 2.0000
```

## CHISQ.DIST

`CHISQ.DIST(x, degrees_freedom, cumulative)`

Returns the chi-squared distribution.

**Example:**
```
=CHISQ.DIST(0.5, 1, TRUE)    → 0.5205
```

## CHISQ.INV

`CHISQ.INV(probability, degrees_freedom)`

Returns the inverse of the left-tailed chi-squared distribution.

**Example:**
```
=CHISQ.INV(0.93, 1)    → 3.2831
```

## CHISQ.DIST.RT

`CHISQ.DIST.RT(x, degrees_freedom)`

Returns the right-tailed probability of the chi-squared distribution.

**Example:**
```
=CHISQ.DIST.RT(18.307, 10)    → 0.0500
```

## CHISQ.INV.RT

`CHISQ.INV.RT(probability, degrees_freedom)`

Returns the inverse of the right-tailed chi-squared distribution.

**Example:**
```
=CHISQ.INV.RT(0.05, 10)    → 18.3070
```

## T.DIST

`T.DIST(x, degrees_freedom, cumulative)`

Returns the Student's t-distribution.

**Example:**
```
=T.DIST(1.0, 5, TRUE)    → 0.8184
```

## T.INV

`T.INV(probability, degrees_freedom)`

Returns the left-tailed inverse of the Student's t-distribution.

**Example:**
```
=T.INV(0.8184, 5)    → 1.0000
```

## T.DIST.2T

`T.DIST.2T(x, degrees_freedom)`

Returns the two-tailed Student's t-distribution.

**Example:**
```
=T.DIST.2T(1.96, 60)    → 0.0546
```

## T.INV.2T

`T.INV.2T(probability, degrees_freedom)`

Returns the two-tailed inverse of the Student's t-distribution.

**Example:**
```
=T.INV.2T(0.0546, 60)    → 1.9600
```

## T.DIST.RT

`T.DIST.RT(x, degrees_freedom)`

Returns the right-tailed Student's t-distribution.

**Example:**
```
=T.DIST.RT(1.0, 5)    → 0.1816
```

## F.DIST

`F.DIST(x, deg_freedom1, deg_freedom2, cumulative)`

Returns the F probability distribution.

**Example:**
```
=F.DIST(15.2069, 6, 4, TRUE)    → 0.99
```

## F.INV

`F.INV(probability, deg_freedom1, deg_freedom2)`

Returns the inverse of the F probability distribution.

**Example:**
```
=F.INV(0.99, 6, 4)    → 15.2069
```

## F.DIST.RT

`F.DIST.RT(x, deg_freedom1, deg_freedom2)`

Returns the right-tailed F probability distribution.

**Example:**
```
=F.DIST.RT(15.2069, 6, 4)    → 0.01
```

## F.INV.RT

`F.INV.RT(probability, deg_freedom1, deg_freedom2)`

Returns the inverse of the right-tailed F probability distribution.

**Example:**
```
=F.INV.RT(0.01, 6, 4)    → 15.2069
```

## WEIBULL.DIST

`WEIBULL.DIST(x, alpha, beta, cumulative)`

Returns the Weibull distribution.

**Example:**
```
=WEIBULL.DIST(105, 20, 100, TRUE)    → 0.9295
```

## LOGNORM.DIST

`LOGNORM.DIST(x, mean, standard_dev, cumulative)`

Returns the lognormal distribution.

**Example:**
```
=LOGNORM.DIST(4, 3.5, 1.2, TRUE)    → 0.0390
```

## LOGNORM.INV

`LOGNORM.INV(probability, mean, standard_dev)`

Returns the inverse of the lognormal cumulative distribution.

**Example:**
```
=LOGNORM.INV(0.0390, 3.5, 1.2)    → 4.0000
```

## CONFIDENCE.NORM

`CONFIDENCE.NORM(alpha, standard_dev, size)`

Returns the confidence interval for a population mean using a normal distribution.

**Example:**
```
=CONFIDENCE.NORM(0.05, 2.5, 50)    → 0.6930
```

## CONFIDENCE.T

`CONFIDENCE.T(alpha, standard_dev, size)`

Returns the confidence interval for a population mean using a Student's t-distribution.

**Example:**
```
=CONFIDENCE.T(0.05, 1, 50)    → 0.2842
```

## Z.TEST

`Z.TEST(array, x, [sigma])`

Returns the one-tailed p-value of a z-test.

**Example:**
```
=Z.TEST({3,6,7,8,6,5,4,2,1,9}, 4)    → 0.0907
```

## T.TEST

`T.TEST(array1, array2, tails, type)`

Returns the probability associated with a Student's t-test.

**Example:**
```
=T.TEST({3,4,5,8,9,1,2,4,5}, {6,19,3,2,14,4,5,17,1}, 2, 1)    → 0.1961
```

## CHISQ.TEST

`CHISQ.TEST(actual_range, expected_range)`

Returns the chi-squared test for independence.

**Example:**
```
=CHISQ.TEST({58,35,11}, {45.35,50.23,8.42})    → 0.0003
```

## F.TEST

`F.TEST(array1, array2)`

Returns the result of an F-test (two-tailed p-value comparing variances).

**Example:**
```
=F.TEST({6,7,9,15,21}, {20,28,31,38,40})    → 0.6483
```

## LINEST

`LINEST(known_ys, [known_xs], [const], [stats])`

Returns statistics for a linear regression line.

**Example:**
```
=LINEST({1,9,5,7}, {0,4,2,3})    → {2, 1}
```

## LOGEST

`LOGEST(known_ys, [known_xs], [const], [stats])`

Returns statistics for an exponential regression curve.

**Example:**
```
=LOGEST({1,5,10,50}, {1,2,3,4})    → {3.4507, 0.4742}
```

## TREND

`TREND(known_ys, [known_xs], [new_xs], [const])`

Returns values along a linear trend.

**Example:**
```
=TREND({1,9,5,7}, {0,4,2,3}, {5})    → {11}
```

## GROWTH

`GROWTH(known_ys, [known_xs], [new_xs], [const])`

Returns values along an exponential growth trend.

**Example:**
```
=GROWTH({33100,47300,69000,102000,150000,220000}, {11,12,13,14,15,16}, {17,18,19})    → {320197, 468536, 685648}
```
