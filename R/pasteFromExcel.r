# pasteFromExcel.r
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

pasteFromExcel <- function(sep = "\t", check.names = FALSE, 
  stringsAsFactors = default.stringsAsFactors(), 
  header = FALSE, rowheader = FALSE, 
  convertFormattedNumbers = TRUE, 
  simplify = TRUE, drop = TRUE, 
  na.strings=c("NA", "", "#DIV/0!"), zero.strings = "-", ...) {
  x <- switch(Sys.info()["sysname"],
    Windows = read.delim(file = "clipboard", sep = sep, header = header, check.names = check.names, stringsAsFactors = stringsAsFactors, na.strings = na.strings, ...),
    Darwin = suppressWarnings(read.delim(file = pipe("pbpaste"), sep = sep, header = header, check.names = check.names, stringsAsFactors = stringsAsFactors, na.strings = na.strings, ...)),
    stop("unsupported OS")
    )
  if (rowheader) {
    row.names(x) <- x[[1]]
    x <- x[-1]
    }
  if (convertFormattedNumbers) for (i in seq_along(x)) {
    if (w <- is.numeric.string(x[[i]], zero.strings = zero.strings)) x[[i]] <- attr(w, "value")
    }
  for (i in seq_along(x)) {
    if (w <- is.date.string(x[[i]])) x[[i]] <- attr(w, "value")
    }
  # 6-6-2013
  #if (simplify) return(.simplifyDataFrame(x, header = FALSE, rowheader = FALSE, drop))
  if (simplify) {
    # 1-6-2014 If "NA" in header and user removed "NA" from na.strings,
    # want result to have "NA" in the column names. But as.matrix.data.frame (in .simplifyDataFrame)
    # automatically converts those to <NA> because that's how read.delim treats them.
    nms <- names(x)
    if (!("NA" %in% na.strings)) {
      ina <- is.na(names(x))
      if (any(ina)) names(x)[ina] <- "NA"
      }
    x <- .simplifyDataFrame(x, drop = drop)
    }
  x
  }

