# writeToCsv.r
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
    
writeToCsv <- function(data, file = deparse(substitute(data)), 
                       header, rowheader = TRUE, 
                       overwrite = FALSE, na = "", ...) {
  maxOverwriteVersion <- 99
  if (!missing(header))
    if (!header) stop("R requires headers for all csv files. Attempt to set header = FALSE invalid.")
  overwrite <- match.arg(as.character(overwrite), c("TRUE", "FALSE", "version"))
#  require(tools, quietly = TRUE)
  fe <- file_ext(file)
  if (fe == "") file <- paste(file, "csv", sep = ".")
  else
  if (fe != "csv") warn <- warning("data written to non-csv extension")
  if (file.exists(file)) {
    if (overwrite == FALSE) stop("File '", file, "' already exists, not overwritten.", sep = "")
    else
    if (overwrite == "version") {
      foundi <- FALSE
      originalFile <- file
      for (i in 1:maxOverwriteVersion) { # look for first nonexistent incremental name
        file <- paste(originalFile, sprintf("%02.0f", i), ".csv", sep = "-")
        foundi <- !file.exists(file)
        if (foundi) break
        }
      if (!foundi) stop("max increment ", maxOverwriteVersion, 
                        " exceeded for file '", file, "'", sep = "")
      }
    }
  # Write data to the csv file
  outdata <- as.data.frame(data)
  if (is.vector(data)) {
    names(outdata) <- deparse(substitute(data))
    if (missing(rowheader)) rowheader <- !is.null(names(data))
    }
  if (rowheader) {
    outdata <- cbind(rownames(outdata), outdata)
    names(outdata)[1] <- ifelse(is.null(names(dimnames(data))[1]), "", names(dimnames(data))[1])
    }
  # If unsuccessful, res will be an "error".
  res <- tryCatch(write.csv(outdata, file = file, row.names = FALSE, na = na, ...), error = function(e) e)
  if ("error" %in% class(res)) warning("Save unsuccessful. File currently open? Invalid filename?")
  else file_path_as_absolute(file)
  }

