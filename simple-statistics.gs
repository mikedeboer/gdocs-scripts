/**
 * The methods defined below were cherry-picked and taken from the 'simple-statistics'
 * module.
 * https://github.com/simple-statistics/simple-statistics
 * ISC License, 2014, Tom MacWright.
 */

/**
 * Our default sum is the [Kahan-Babuska algorithm](https://pdfs.semanticscholar.org/1760/7d467cda1d0277ad272deb2113533131dc09.pdf).
 * This method is an improvement over the classical
 * [Kahan summation algorithm](https://en.wikipedia.org/wiki/Kahan_summation_algorithm).
 * It aims at computing the sum of a list of numbers while correcting for
 * floating-point errors. Traditionally, sums are calculated as many
 * successive additions, each one with its own floating-point roundoff. These
 * losses in precision add up as the number of numbers increases. This alternative
 * algorithm is more accurate than the simple way of calculating sums by simple
 * addition.
 *
 * This runs on `O(n)`, linear time in respect to the array.
 *
 * @param {Array<number>} x input
 * @return {number} sum of all input numbers
 * @example
 * sum([1, 2, 3]); // => 6
 */
function sum(x) {
  // If the array is empty, we needn't bother computing its sum
  if (x.length === 0) {
    return 0;
  }

  // Initializing the sum as the first number in the array
  var sum = x[0];

  // Keeping track of the floating-point error correction
  var correction = 0;

  var transition;

  for (var i = 1; i < x.length; i++) {
    transition = sum + x[i];

    // Here we need to update the correction in a different fashion
    // if the new absolute value is greater than the absolute sum
    if (Math.abs(sum) >= Math.abs(x[i])) {
      correction += sum - transition + x[i];
    } else {
      correction += x[i] - transition + sum;
    }

    sum = transition;
  }

  // Returning the corrected sum
  return sum + correction;
}

/**
 * The mean, _also known as average_,
 * is the sum of all values over the number of values.
 * This is a [measure of central tendency](https://en.wikipedia.org/wiki/Central_tendency):
 * a method of finding a typical or central value of a set of numbers.
 *
 * This runs on `O(n)`, linear time in respect to the array
 *
 * @param {Array<number>} x sample of one or more data points
 * @throws {Error} if the the length of x is less than one
 * @returns {number} mean
 * @example
 * mean([0, 10]); // => 5
 */
function mean(x) {
  // The mean of no numbers is null
  if (x.length === 0) {
    throw new Error("mean requires at least one data point");
  }

  return sum(x) / x.length;
}

/**
 * The sum of deviations to the Nth power.
 * When n=2 it's the sum of squared deviations.
 * When n=3 it's the sum of cubed deviations.
 *
 * @param {Array<number>} x
 * @param {number} n power
 * @returns {number} sum of nth power deviations
 *
 * @example
 * var input = [1, 2, 3];
 * // since the variance of a set is the mean squared
 * // deviations, we can calculate that with sumNthPowerDeviations:
 * sumNthPowerDeviations(input, 2) / input.length;
 */
function sumNthPowerDeviations(x, n) {
  var meanValue = mean(x);
  var sum = 0;
  var tempValue;
  var i;

  // This is an optimization: when n is 2 (we're computing a number squared),
  // multiplying the number by itself is significantly faster than using
  // the Math.pow method.
  if (n === 2) {
    for (i = 0; i < x.length; i++) {
      tempValue = x[i] - meanValue;
      sum += tempValue * tempValue;
    }
  } else {
    for (i = 0; i < x.length; i++) {
      sum += Math.pow(x[i] - meanValue, n);
    }
  }

  return sum;
}

/**
 * The [variance](http://en.wikipedia.org/wiki/Variance)
 * is the sum of squared deviations from the mean.
 *
 * This is an implementation of variance, not sample variance:
 * see the `sampleVariance` method if you want a sample measure.
 *
 * @param {Array<number>} x a population of one or more data points
 * @returns {number} variance: a value greater than or equal to zero.
 * zero indicates that all values are identical.
 * @throws {Error} if x's length is 0
 * @example
 * variance([1, 2, 3, 4, 5, 6]); // => 2.9166666666666665
 */
function variance(x) {
  // The variance of no numbers is null
  if (x.length === 0) {
    throw new Error("variance requires at least one data point");
  }

  // Find the mean of squared deviations between the
  // mean value and each value.
  return sumNthPowerDeviations(x, 2) / x.length;
}

/**
 * The [standard deviation](http://en.wikipedia.org/wiki/Standard_deviation)
 * is the square root of the variance. This is also known as the population
 * standard deviation. It's useful for measuring the amount
 * of variation or dispersion in a set of values.
 *
 * Standard deviation is only appropriate for full-population knowledge: for
 * samples of a population, {@link sampleStandardDeviation} is
 * more appropriate.
 *
 * @param {Array<number>} x input
 * @returns {number} standard deviation
 * @example
 * variance([2, 4, 4, 4, 5, 5, 7, 9]); // => 4
 * standardDeviation([2, 4, 4, 4, 5, 5, 7, 9]); // => 2
 */
function standardDeviation(x) {
  if (x.length === 1) {
    return 0;
  }
  var v = variance(x);
  return Math.sqrt(v);
}

/**
 * The [quantile](https://en.wikipedia.org/wiki/Quantile):
 * this is a population quantile, since we assume to know the entire
 * dataset in this library. This is an implementation of the
 * [Quantiles of a Population](http://en.wikipedia.org/wiki/Quantile#Quantiles_of_a_population)
 * algorithm from wikipedia.
 *
 * Sample is a one-dimensional array of numbers,
 * and p is either a decimal number from 0 to 1 or an array of decimal
 * numbers from 0 to 1.
 * In terms of a k/q quantile, p = k/q - it's just dealing with fractions or dealing
 * with decimal values.
 * When p is an array, the result of the function is also an array containing the appropriate
 * quantiles in input order
 *
 * @param {Array<number>} x sample of one or more numbers
 * @param {Array<number> | number} p the desired quantile, as a number between 0 and 1
 * @returns {number} quantile
 * @example
 * quantile([3, 6, 7, 8, 8, 9, 10, 13, 15, 16, 20], 0.5); // => 9
 */
function quantile(x, p) {
  var copy = x.slice();

  if (Array.isArray(p)) {
    // rearrange elements so that each element corresponding to a requested
    // quantile is on a place it would be if the array was fully sorted
    multiQuantileSelect(copy, p);
    // Initialize the result array
    var results = [];
    // For each requested quantile
    for (var i = 0; i < p.length; i++) {
      results[i] = quantileSorted(copy, p[i]);
    }
    return results;
  }

  var idx = quantileIndex(copy.length, p);
  quantileSelect(copy, idx, 0, copy.length - 1);
  return quantileSorted(copy, p);
}

function quantileSelect(arr, k, left, right) {
  if (k % 1 === 0) {
    quickselect(arr, k, left, right);
  } else {
    k = Math.floor(k);
    quickselect(arr, k, left, right);
    quickselect(arr, k + 1, k + 1, right);
  }
}

function multiQuantileSelect(arr, p) {
  var indices = [0];
  for (var i = 0; i < p.length; i++) {
    indices.push(quantileIndex(arr.length, p[i]));
  }
  indices.push(arr.length - 1);
  indices.sort(compare);

  var stack = [0, indices.length - 1];
  var r, l, m;
  while (stack.length) {
    r = Math.ceil(stack.pop());
    l = Math.floor(stack.pop());
    if (r - l <= 1) {
      continue;
    }

    m = Math.floor((l + r) / 2);
    quantileSelect(arr, indices[m], Math.floor(indices[l]), Math.ceil(indices[r]));

    stack.push(l, m, m, r);
  }
}

function compare(a, b) {
  return a - b;
}

function quantileIndex(len, p) {
  var idx = len * p;
  if (p === 1) {
    // If p is 1, directly return the last index
    return len - 1;
  } else if (p === 0) {
    // If p is 0, directly return the first index
    return 0;
  } else if (idx % 1 !== 0) {
    // If index is not integer, return the next index in array
    return Math.ceil(idx) - 1;
  } else if (len % 2 === 0) {
    // If the list has even-length, we'll return the middle of two indices
    // around quantile to indicate that we need an average value of the two
    return idx - 0.5;
  }

  // Finally, in the simple case of an integer index
  // with an odd-length list, return the index
  return idx;
}

/**
 * This is the internal implementation of quantiles: when you know
 * that the order is sorted, you don't need to re-sort it, and the computations
 * are faster.
 *
 * @param {Array<number>} x sample of one or more data points
 * @param {number} p desired quantile: a number between 0 to 1, inclusive
 * @returns {number} quantile value
 * @throws {Error} if p ix outside of the range from 0 to 1
 * @throws {Error} if x is empty
 * @example
 * quantileSorted([3, 6, 7, 8, 8, 9, 10, 13, 15, 16, 20], 0.5); // => 9
 */
function quantileSorted(x, p) {
  var idx = x.length * p;
  if (x.length === 0) {
    throw new Error("quantile requires at least one data point.");
  } else if (p < 0 || p > 1) {
    throw new Error("quantiles must be between 0 and 1");
  } else if (p === 1) {
    // If p is 1, directly return the last element.
    return x[x.length - 1];
  } else if (p === 0) {
    // If p is 0, directly return the first element.
    return x[0];
  } else if (idx % 1 !== 0) {
    // If p is not integer, return the next element in array.
    return x[Math.ceil(idx) - 1];
  } else if (x.length % 2 === 0) {
    // If the list has even-length, we'll take the average of this number
    // and the next value, if there is one.
    return (x[idx - 1] + x[idx]) / 2;
  }

  // Finally, in the simple case of an integer value
  // with an odd-length list, return the x value at the index.
  return x[idx];
}

/**
 * Rearrange items in `arr` so that all items in `[left, k]` range are the smallest.
 * The `k`-th element will have the `(k - left + 1)`-th smallest value in `[left, right]`.
 *
 * Implements Floyd-Rivest selection algorithm https://en.wikipedia.org/wiki/Floyd-Rivest_algorithm
 *
 * @param {Array<number>} arr input array
 * @param {number} k pivot index
 * @param {number} [left] left index
 * @param {number} [right] right index
 * @returns {void} mutates input array
 * @example
 * var arr = [65, 28, 59, 33, 21, 56, 22, 95, 50, 12, 90, 53, 28, 77, 39];
 * quickselect(arr, 8);
 * // = [39, 28, 28, 33, 21, 12, 22, 50, 53, 56, 59, 65, 90, 77, 95]
 */
function quickselect(arr, k, left, right) {
  left = left || 0;
  right = right || arr.length - 1;

  var n, m, z, s, sd, newLeft, newRight, t, i, j;
  while (right > left) {
    // 600 and 0.5 are arbitrary constants chosen in the original paper to minimize execution time
    if (right - left > 600) {
      n = right - left + 1;
      m = k - left + 1;
      z = Math.log(n);
      s = 0.5 * Math.exp((2 * z) / 3);
      sd = 0.5 * Math.sqrt((z * s * (n - s)) / n);
      if (m - n / 2 < 0)  {
        sd *= -1;
      }
      newLeft = Math.max(left, Math.floor(k - (m * s) / n + sd));
      newRight = Math.min(right, Math.floor(k + ((n - m) * s) / n + sd));
      quickselect(arr, k, newLeft, newRight);
    }

    t = arr[k];
    i = left;
    j = right;

    swap(arr, left, k);
    if (arr[right] > t) {
      swap(arr, left, right);
    }

    while (i < j) {
      swap(arr, i, j);
      i++;
      j--;
      while (arr[i] < t) {
        i++;
      }
      while (arr[j] > t) {
        j--;
      }
    }

    if (arr[left] === t) {
      swap(arr, left, j);
    } else {
      j++;
      swap(arr, j, right);
    }

    if (j <= k) {
      left = j + 1;
    }
    if (k <= j) {
      right = j - 1;
    }
  }
}

function swap(arr, i, j) {
  var tmp = arr[i];
  arr[i] = arr[j];
  arr[j] = tmp;
}

/**
 * The [Interquartile range](http://en.wikipedia.org/wiki/Interquartile_range) is
 * a measure of statistical dispersion, or how scattered, spread, or
 * concentrated a distribution is. It's computed as the difference between
 * the third quartile and first quartile.
 *
 * @param {Array<number>} x sample of one or more numbers
 * @returns {number} interquartile range: the span between lower and upper quartile,
 * 0.25 and 0.75
 * @example
 * interquartileRange([0, 1, 2, 3]); // => 2
 */
function interquartileRange(x) {
  // Interquartile range is the span between the upper quartile,
  // at `0.75`, and lower quartile, `0.25`
  return quantile(x, 0.75) - quantile(x, 0.25);
}

/**
 * The [Z-Score, or Standard Score](http://en.wikipedia.org/wiki/Standard_score).
 *
 * The standard score is the number of standard deviations an observation
 * or datum is above or below the mean. Thus, a positive standard score
 * represents a datum above the mean, while a negative standard score
 * represents a datum below the mean. It is a dimensionless quantity
 * obtained by subtracting the population mean from an individual raw
 * score and then dividing the difference by the population standard
 * deviation.
 *
 * The z-score is only defined if one knows the population parameters;
 * if one only has a sample set, then the analogous computation with
 * sample mean and sample standard deviation yields the
 * Student's t-statistic.
 *
 * @param {number} x
 * @param {number} mean
 * @param {number} standardDeviation
 * @return {number} z score
 * @example
 * zScore(78, 80, 5); // => -0.4
 */
function zScore(x, mean, standardDeviation) {
  return (x - mean) / standardDeviation;
}

function pruneOutliers(x) {
  var work = [].concat(x);
  work.sort(compare);
  var len = x.length;
  if (len <= 2) {
    return work;
  }

  var res = [];
  var m = mean(work);
  var stdev = standardDeviation(work);

  for (var i = 0; i < len; ++i) {
    // if (work[i] >= (q1 - 1.5 * iqr) && work[i] <= (q3 + 1.5 * iqr))
    // Yeah, magic constant *sigh*.
    if (zScore(work[i], m, stdev) <= -0.9) {
      continue;
    }
    res.push(work[i]);
  }
  return res;
}
