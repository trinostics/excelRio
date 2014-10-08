# readFromCsv.r
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
  
readFromCsv <- function(file, 
  stringsAsFactors = default.stringsAsFactors(),
  simplify = TRUE, drop = TRUE, 
  na.strings = c("NA","#DIV/0!"),
  zero.strings = "-",  
  convertFormattedNumbers = TRUE, ..., header = TRUE, rowheader = FALSE) {
  if (missing(file)) {
    file <- choose.files()
    if (length(file) == 0) # hit Cancel
      return(file)
    }
  if (!file.exists(file)) {
    warning("File '", file, "' does not exist.", sep = "")
    return(NULL)
    }
#  require(tools, quietly = TRUE)
  ext <- file_ext(file)
  if (tolower(ext) != "csv") {
    warning("Only csv files can be read with readFromCsv.")
    return(character(0))
    }
  if (rowheader) row.names = 1
  else row.names = NULL
  x <- read.csv(file, header = header, row.names = row.names, stringsAsFactors = stringsAsFactors, na.strings = na.strings, ...)
  if (convertFormattedNumbers) for (i in seq_along(x)) {
    if (w <- is.numeric.string(x[[i]], zero.strings = zero.strings)) x[[i]] <- attr(w, "value")
    }
  for (i in seq_along(x)) {
    if (w <- is.date.string(x[[i]])) x[[i]] <- attr(w, "value")
    }
  if (simplify) x <- .simplifyDataFrame(x, drop = drop)

  attr(x, "basename") <- basename(file)
  attr(x, "dirname") <- dirname(file)
  attr(x, "fullname") <- file.path(attr(x, "dirname"), attr(x, "basename"))
  attr(x, "file_ext") <- tools::file_ext(file)
  x
  } 

