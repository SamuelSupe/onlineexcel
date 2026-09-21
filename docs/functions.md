# Formula capabilities

187 distinct functions. Optional parameters appear in brackets. Functions use invariant English names and comma separators. This list describes implemented functions, not certification against every Excel edge case.

## array

| Function | Parameters | Notes |
| --- | --- | --- |
| CHOOSECOLS | array, col_num1, ... |  |
| CHOOSEROWS | array, row_num1, ... |  |
| DROP | array, rows, [columns] |  |
| FILTER | array, include, [if_empty] |  |
| HSTACK | array1, ... |  |
| SEQUENCE | rows, [columns], [start], [step] |  |
| SORT | array, [sort_index], [sort_order], [by_col] |  |
| SORTBY | array, by_array1, [sort_order1], ... | Sort keys are vertical columns. |
| TAKE | array, rows, [columns] |  |
| TRANSPOSE | array |  |
| UNIQUE | array, [by_col], [exactly_once] |  |
| VSTACK | array1, ... |  |

## date

| Function | Parameters | Notes |
| --- | --- | --- |
| DATE | year, month, day |  |
| DATEVALUE | value | Accepts ISO YYYY-MM-DD; ambiguous locale-dependent strings are rejected. |
| DAY | value |  |
| DAYS | end_date, start_date |  |
| DAYS360 | start_date, end_date, [method] |  |
| EDATE | start_date, months |  |
| EOMONTH | start_date, months |  |
| HOUR | value |  |
| ISOWEEKNUM | value |  |
| MINUTE | value |  |
| MONTH | value |  |
| NETWORKDAYS | start_date, end_date, [holidays] |  |
| NOW |  |  |
| SECOND | value |  |
| TIME | hour, minute, second |  |
| TIMEVALUE | value |  |
| TODAY |  |  |
| WEEKDAY | serial_number, [return_type] |  |
| WEEKNUM | serial_number, [return_type] |  |
| WORKDAY | start_date, days, [holidays] |  |
| YEAR | value |  |

## financial

| Function | Parameters | Notes |
| --- | --- | --- |
| EFFECT | nominal_rate, npery |  |
| FV | rate, nper, pmt, [pv], [type] |  |
| IRR | values, [guess] |  |
| NOMINAL | effect_rate, npery |  |
| NPER | rate, pmt, pv, [fv], [type] |  |
| NPV | rate, value1, ... |  |
| PMT | rate, nper, pv, [fv], [type] |  |
| PV | rate, nper, pmt, [fv], [type] |  |
| RATE | nper, pmt, pv, [fv], [type], [guess] |  |
| SLN | cost, salvage, life |  |
| SYD | cost, salvage, life, per |  |
| XIRR | values, dates, [guess] |  |
| XNPV | rate, values, dates |  |

## information

| Function | Parameters | Notes |
| --- | --- | --- |
| ISBLANK | value |  |
| ISERR | value |  |
| ISERROR | value |  |
| ISEVEN | value |  |
| ISLOGICAL | value |  |
| ISNA | value |  |
| ISNONTEXT | value |  |
| ISNUMBER | value |  |
| ISODD | value |  |
| ISTEXT | value |  |
| N | value |  |
| NA |  |  |
| T | value |  |
| TYPE | value |  |

## logical

| Function | Parameters | Notes |
| --- | --- | --- |
| AND | value1, ... |  |
| CHOOSE | index_num, value1, ... |  |
| FALSE |  |  |
| IF | logical_test, value_if_true, [value_if_false] |  |
| IFERROR | value, value_if_error |  |
| IFNA | value, value_if_na |  |
| IFS | logical_test1, value1, ... |  |
| NOT | value |  |
| OR | value1, ... |  |
| SWITCH | expression, value1, result1, ..., [default] |  |
| TRUE |  |  |
| XOR | value1, ... |  |

## lookup

| Function | Parameters | Notes |
| --- | --- | --- |
| ADDRESS | row_num, column_num, [abs_num], [a1], [sheet_text] |  |
| COLUMN | [reference] |  |
| COLUMNS | array |  |
| HLOOKUP | lookup_value, table_array, row_index_num, [range_lookup] |  |
| INDEX | array, row_num, [column_num] |  |
| MATCH | lookup_value, lookup_array, [match_type] |  |
| ROW | [reference] |  |
| ROWS | array |  |
| VLOOKUP | lookup_value, table_array, col_index_num, [range_lookup] |  |
| XLOOKUP | lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode] |  |
| XMATCH | lookup_value, lookup_array, [match_mode], [search_mode] |  |

## math

| Function | Parameters | Notes |
| --- | --- | --- |
| ABS | value |  |
| ACOS | value |  |
| ACOSH | value |  |
| ASIN | value |  |
| ASINH | value |  |
| ATAN | value |  |
| ATAN2 | x_num, y_num |  |
| ATANH | value |  |
| CEILING.MATH | number, [significance], [mode] |  |
| COMBIN | number, number_chosen |  |
| COS | value |  |
| COSH | value |  |
| DEGREES | value |  |
| EVEN | value |  |
| EXP | value |  |
| FACT | value |  |
| FLOOR.MATH | number, [significance], [mode] |  |
| GCD | value1, ... |  |
| INT | value |  |
| LCM | value1, ... |  |
| LN | value |  |
| LOG | number, [base] |  |
| LOG10 | value |  |
| MOD | number, divisor |  |
| MROUND | number, multiple |  |
| ODD | value |  |
| PI |  |  |
| POWER | number, power |  |
| QUOTIENT | numerator, denominator |  |
| RADIANS | value |  |
| RAND |  |  |
| RANDBETWEEN | bottom, top |  |
| ROUND | number, num_digits |  |
| ROUNDDOWN | number, num_digits |  |
| ROUNDUP | number, num_digits |  |
| SIGN | value |  |
| SIN | value |  |
| SINH | value |  |
| SQRT | value |  |
| SUMPRODUCT | array1, [array2], ... |  |
| TAN | value |  |
| TANH | value |  |
| TRUNC | number, [num_digits] |  |

## statistics

| Function | Parameters | Notes |
| --- | --- | --- |
| AVEDEV | value1, ... |  |
| AVERAGE | value1, ... |  |
| AVERAGEIF | range, criteria, [average_range] |  |
| AVERAGEIFS | average_range, criteria_range1, criteria1, ... |  |
| CORREL | array1, array2 |  |
| COUNT | value1, ... |  |
| COUNTA | value1, ... |  |
| COUNTBLANK | value |  |
| COUNTIF | range, criteria |  |
| COUNTIFS | criteria_range1, criteria1, ... |  |
| COVARIANCE.P | array1, array2 |  |
| COVARIANCE.S | array1, array2 |  |
| DEVSQ | value1, ... |  |
| GEOMEAN | value1, ... |  |
| HARMEAN | value1, ... |  |
| INTERCEPT | known_ys, known_xs |  |
| LARGE | array, k |  |
| MAX | value1, ... |  |
| MEDIAN | value1, ... |  |
| MIN | value1, ... |  |
| MODE.SNGL | value1, ... |  |
| PERCENTILE.INC | array, k |  |
| PRODUCT | value1, ... |  |
| QUARTILE.INC | array, quart |  |
| RANK.EQ | number, ref, [order] |  |
| RSQ | known_ys, known_xs |  |
| SLOPE | known_ys, known_xs |  |
| SMALL | array, k |  |
| STDEV.P | value1, ... |  |
| STDEV.S | value1, ... |  |
| SUM | value1, ... |  |
| SUMIF | range, criteria, [sum_range] |  |
| SUMIFS | sum_range, criteria_range1, criteria1, ... |  |
| SUMSQ | value1, ... |  |
| VAR.P | value1, ... |  |
| VAR.S | value1, ... |  |

## text

| Function | Parameters | Notes |
| --- | --- | --- |
| CHAR | value | Uses Unicode code points 1–255; platform-specific legacy code pages are not emulated. |
| CLEAN | value |  |
| CODE | value |  |
| CONCAT | text1, ... |  |
| EXACT | text1, text2 |  |
| FIND | find_text, within_text, [start_num] |  |
| FIXED | number, [decimals], [no_commas] |  |
| LEFT | text, [num_chars] |  |
| LEN | value |  |
| LOWER | value |  |
| MID | text, start_num, num_chars |  |
| NUMBERVALUE | text, [decimal_separator], [group_separator] |  |
| PROPER | value |  |
| REPLACE | old_text, start_num, num_chars, new_text |  |
| REPT | text, number_times |  |
| RIGHT | text, [num_chars] |  |
| SEARCH | find_text, within_text, [start_num] |  |
| SUBSTITUTE | text, old_text, new_text, [instance_num] |  |
| TEXT | value, format_text |  |
| TEXTJOIN | delimiter, ignore_empty, text1, ... |  |
| TRIM | value |  |
| UNICHAR | value |  |
| UNICODE | value |  |
| UPPER | value |  |
| VALUE | value | Invariant decimal syntax; locale-specific currency and date strings are not parsed. |

