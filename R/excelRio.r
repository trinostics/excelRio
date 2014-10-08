# excelRio.r
##    Copyright (C) <2012 - 2014>  <Daniel Murphy>

#    This file is part of excelRio.
#
#    excelRio is free software: you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation, either version 3 of the License, or
#    (at your option) any later version.
#
#    excelRio is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#
#    You should have received a copy of the GNU General Public License
#    along with excelRio.  If not, see <http://www.gnu.org/licenses/>.
    
## 7/16/2014
## XLConnect/rJava problems after Java is updated.
##    If you get the error "JAVA_HOME cannot be determined from the Registry" 
## Per Simon Urbanek at
## http://www.r-statistics.com/2012/08/how-to-load-the-rjava-package-after-the-error-java_home-cannot-be-determined-from-the-registry/
##    could be 32- 64-bit issue, which just happened to me -- Java notified me that an update was available,
##    I followed all the steps to successfully install Java (which was probably the 32-bit version),
#     then got the error above when I tried loading rJava in 64-bit R.
## Goto http://www.java.com/en/download/manual.jsp and download the version
##    of Java you want. I did that, installed the 64-bit version, then
#     rJava loaded without the error.

# From: http://stackoverflow.com/questions/354044/what-is-the-best-u-s-currency-regex
# currencypatternA is the original pattern from the above website, the one that said (and got one vote):
#   I found this regular expression on line at www.RegExLib.com by Kirk Fuller, Gregg Durishan
#   I've been using it successfully for the past couple of years.
# $ sign inside ( being ok -- not Excel approach
currencypatternA  <- "^\\$?\\-?([1-9]{1}[0-9]{0,2}(\\,\\d{3})*(\\.\\d{0,2})?|[1-9]{1}\\d{0,}(\\.\\d{0,2})?|0(\\.\\d{0,2})?|(\\.\\d{1,2}))$|^\\-?\\$?([1-9]{1}\\d{0,2}(\\,\\d{3})*(\\.\\d{0,2})?|[1-9]{1}\\d{0,}(\\.\\d{0,2})?|0(\\.\\d{0,2})?|(\\.\\d{1,2}))$|^\\(\\$?([1-9]{1}\\d{0,2}(\\,\\d{3})*(\\.\\d{0,2})?|[1-9]{1}\\d{0,}(\\.\\d{0,2})?|0(\\.\\d{0,2})?|(\\.\\d{1,2}))\\)$"
# $ sign before( is ok -- Excel approach
currencypattern <- "^\\$?\\-?([1-9]{1}[0-9]{0,2}(\\,\\d{3})*(\\.\\d{0,2})?|[1-9]{1}\\d{0,}(\\.\\d{0,2})?|0(\\.\\d{0,2})?|(\\.\\d{1,2}))$|^\\-?\\$?([1-9]{1}\\d{0,2}(\\,\\d{3})*(\\.\\d{0,2})?|[1-9]{1}\\d{0,}(\\.\\d{0,2})?|0(\\.\\d{0,2})?|(\\.\\d{1,2}))$|^\\$?\\(([1-9]{1}\\d{0,2}(\\,\\d{3})*(\\.\\d{0,2})?|[1-9]{1}\\d{0,}(\\.\\d{0,2})?|0(\\.\\d{0,2})?|(\\.\\d{1,2}))\\)$"
# Here's a test
# x <- c("1,234.00", "12,34.00", "$1,000", "(124)", "$(123)", "($123)", "  1,000   ", "NA")
# grep(currencypattern, trim(x))
# [1] 1 3 4 5 7  # correct answer -- 8th NA element is picked up by default na.strings argument in is.numeric.string

trim <- function(x) sub("[[:space:]]+$", "", sub("^[[:space:]]+", "", x))
is.numeric.string <- function(x, na.strings = c("NA", "#DIV/0!"), inf.strings = c("Inf", "-Inf"), zero.strings = NULL) {
  xs <- trim(x)
  # 2014-05-27: '-' sign only signifying Excel-formatted zero is ok, set by
  #             new zero.strings argument; NULL (default) omits test.
  #             zero.strings argument also added to pasteFromExcel, 
  #             readFromCsv and readFromExcel with default value "-".
  # Here's a test, with correct answers
  # x <- c("1,234.00", "$1,000", "(124)", "$(123)", "  1,000   ", "NA", " - ", "Inf")
  # is.numeric.string(x)
  #[1] FALSE
  # is.numeric.string(x, zero.strings = "-")
  #[1] TRUE
  #attr(,"value")
  #[1] 1234 1000 -124 -123 1000   NA    0  Inf
  xs[xs %in% zero.strings] <- 0
  ok.currencypattern <- (ixs <- seq_along(xs)) %in% grep(currencypattern, xs)
  ok.na <- is.na(xs)
  ok.na.strings <- xs %in% na.strings
  ok.inf.strings <- xs %in% inf.strings
  result <- all(ok.currencypattern | ok.na | ok.na.strings | ok.inf.strings)
  if (result) {
    xs <- gsub('[\\$\\),]', '', xs) # delete most extraneous characters
    is.negative <- ixs %in% grep('^[-\\(]', xs)
    xs <- gsub('^[-\\(]', '', xs) # delete the negative extraneous characters
    xs[ok.na.strings] <- NA
    value <- as.numeric(xs)
    value[is.negative] <- -value[is.negative]
    attr(result, "value") <- value
    }
  result
  }
is.date.string <- function(x, ...) {
  if (!is.character(x) && !is.factor(x)) return(FALSE)
#  require(mondate)
  y <- tryCatch(mondate(x, ...), warning = function(e) e)
  if (is.null(y)) return(FALSE)
  if (class(y)[1] %in% c("simpleWarning","warning", "condition")) return(FALSE)
  structure(TRUE, value = as.Date(y))
  }

#.simplifyDataFrame <- function(x, header, rowheader, drop) {
# 6-6-2013
.simplifyDataFrame <- function(x, drop) {
  # If all columns of a data.frame are the same class, return a matrix
  # If drop and one of the dimensions of the matrix has length 1, return a vector
  cls <- sapply(x, function(y) class(y)[1L])
  if (any(cls != cls[1L])) return(x)
  if (all(cls == "factor")){ # assume want to keep factor representation of columns
    if (length(x) == 1L) return(x[[1L]]) # can return a factor vector 
    # Can't represent a factor in more than 1 dimension even if all levels
    # of all columns are the same, so just return x
    return(x)
    }
  as.matrix(x)[, , drop = drop]
  }


