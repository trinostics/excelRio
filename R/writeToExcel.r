# writeToExcel.r
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
    
## 9/8/2014:
# writeToExcel can now use WriteXLS package
# Pros
#   1) it can handle larger files than XLConnect
#   2) does not need java!
#   3) it prints out dates with day-only formatting (no time)
#   4) it also prints NA's as blanks
#
# Cons
#   1) it overwrites by default
#   2) it cannot iterate another sheet, so cannot add data to an existing workbook
#   3) it cannot start writing data to an upper left corner other than A1
# Due to the Cons, XLConnect is still the default.

writeToExcel <- function(data, file = deparse(substitute(data)), 
                         sheet = "Sheet1", 
                         startRow = 1, startCol = 1, 
                         header = TRUE, rowheader = TRUE, 
                         overwrite = TRUE,
                         pkg = "XLConnect",
                         ...) {
  # Uses the XLConnect package so NA's print as blanks
  # Note: XLConnect OVERWRITES cell contents if file & sheet exist
  # If file doesn't exist, writes to specified sheet, or Sheet1 by default
  if (!missing(sheet)) if (!(is.character(sheet) | length(sheet)>1)) stop("sheet argument must be a character scalar")
  if (length(file) > 1 | !is.character(file)) stop("invalid 'file' argument")
  missing.overwrite <- missing(overwrite)
  overwrite <- match.arg(as.character(overwrite), c("TRUE", "FALSE", "version"))
  if (startRow < 1L) stop("A 'startRow' of ", startRow, " is out of range")
  if (startCol < 1L) stop("A 'startCol' of ", startCol, " is out of range")
  if (length(startRow) > 1L) stop("startRow must have length 1")
  if (length(startCol) > 1L) stop("startCol must have length 1")

#  require(tools, quietly = TRUE)
  if (file_ext(file) == "") file <- paste(file, "xlsx", sep = ".")
  file_exists <- !is.na(file.info(file)$size)

  res <- if (pkg == "XLConnect") {
    maxSheetVersion <- 6
    pkg.exists <- suppressPackageStartupMessages(require(XLConnect, warn.conflicts = FALSE, quietly = TRUE))
    if (!pkg.exists) stop("XLConnect package is not installed. Data not saved")
    if (file_exists) {
      step <- "loading"
      res <- wb <- tryCatch(XLConnect::loadWorkbook(file, create = FALSE), error = function(x) x)
      if (class(res)[1] != "workbook") {
        stop("Excel file '", file_path_as_absolute(file), "' already exists but cannot be loaded. Corrupt?\n", sep = "")
        # return(FALSE) # couldn't load for some reason -- see warning at console
        }
      # Let's see if the sheet exists
      if (missing(sheet)) {
        shts <- getSheets(wb)
        sheet <- paste("Sheet", length(shts)+1, sep = "")
        }
      else
      if (existsSheet(wb, name = sheet)) {
        if (overwrite == "FALSE") {
          res <- simpleError(paste("\a Sheet '", sheet, "' exists in file '", file, "', no overwrite requested. Data not saved.\n", sep = ""))
          return(FALSE)
          }
        else 
        if (overwrite == "iterate") {
          foundi <- FALSE
          originalSheet <- sheet
          for (i in 1:maxSheetVersion) { # look for first nonexistent incremental name
            sheet <- sprintf("%s(%i)", originalSheet, i)
            foundi <- !existsSheet(wb, name = sheet)
            if (foundi) break
            }
          if (!foundi) stop("\a Max sheet version'", sheet, " already exists in file '", file, "'. Data not saved.\n", sep = "")
          }
        }
      }
      .writeToExcel.XLConnect(
        data = data, 
        file = file, 
        sheet = sheet,
        startRow = startRow,
        startCol = startCol,
        header = header,
        rowheader = rowheader)
    }
  else {
    pkg.exists <- suppressPackageStartupMessages(require(WriteXLS, warn.conflicts = FALSE, quietly = TRUE))
    if (!pkg.exists) stop("WriteXLS package is not installed. Data not saved")
    if (!is.character(data)) data <- deparse(substitute(data))
    if (length(data) > 1L) stop("'data' must have length 1")
    if (!inherits(get(data, envir = parent.frame()), "data.frame")) stop(data, " must be a data.frame for WriteXLS")
    if (file_exists) {
      if (missing.overwrite) stop("file '", file, "' already exists. Specify 'overwrite = TRUE' to overwrite")
      else
      if (overwrite != "TRUE") stop("file '", file, "' already exists. Data not saved")
      }
    .writeToExcel.WriteXLS(
      data, 
      file = file, 
      sheet = sheet,
      header = header,
      rowheader = rowheader,
      envir = parent.frame(),
      ...)
    
    }
  if (res) paste(file_path_as_absolute(file), sheet, sep = ":")
  else "Save Failed"
  }

.writeToExcel.XLConnect <- function(data, file, sheet, startRow, startCol, 
                         header, rowheader) {
  # Uses the XLConnect package so NA's print as blanks
  # Note: XLConnect OVERWRITES cell contents if file & sheet exist
  # If file doesn't exist, writes to specified sheet, or Sheet1 by default
  on.exit(if (!is.null(res)) {
    cat("Save failed! See also Java message below. If out of memory, try pkg = WriteXLS, or try writeToCsv instead.\n")
    #    if (step == "saving") cat("file currently open in Excel?\n")
    if (!is.null(res)) warning(res$message)
    })
  res <- wb <- tryCatch(XLConnect::loadWorkbook(file, create = TRUE), error = function(x) x)
  if (class(res)[1] != "workbook") return(FALSE) # couldn't load for some reason -- see warning at console
  res <- tryCatch(XLConnect::createSheet(wb, name = sheet), error = function(x) x)
  if (!is.null(res)) {
    wb <- NULL
    return(FALSE)
    }

  # Don't want XLConnect's default gray background of headers
  # 6/14/2014 the "NONE" action was deleting the styles for Dates
  #  Hmm...setStyleAction was generating changing styles -- sometimes dates were dates
  #     and sometimes dates were numbers -- removed entirely for the time being.
  #  res <- tryCatch((wb, XLC$"STYLE_ACTION.DATA_FORMAT_ONLY"), error = function(x) x)
  #  res <- tryCatch(setStyleAction(wb, XLC$"STYLE_ACTION.NONE"), error = function(x) x)
#  if (!is.null(res)) {
#    wb <- NULL
#    return(FALSE)
#    }

  # Write data to the sheet
  outdata <- as.data.frame(data)
  if (is.vector(data)) {
    names(outdata) <- deparse(substitute(data))
    if (missing(rowheader)) rowheader <- !is.null(names(data))
    }
  # 6/14/2014: Per David Smith in http://blog.revolutionanalytics.com/2009/06/converting-time-zones.html
  #   convert Dates to POSIXct objects via format
  for (i in seq_along(outdata)) if (inherits(outdata[[i]], "Date")) outdata[[i]] <- as.POSIXct(format(outdata[[i]]))

  if (rowheader) {
    outdata <- cbind(rownames(outdata), outdata)
    names(outdata)[1] <- ifelse(is.null(names(dimnames(data))[1]), "", names(dimnames(data))[1])
    }

  # 6/18/2014 Consider using rownames argument ala rownames="Row Names" Message: 27
  #           Date: Tue, 17 Jun 2014 17:11:14 -0400
  #           From: jim holtman <jholtman@gmail.com>
  #           To: R R-help <r-help@stat.math.ethz.ch>
  res <- tryCatch(writeWorksheet(wb, outdata, sheet = sheet, startRow = startRow, startCol = startCol, header = header), error = function(x) x)
  if (!is.null(res)) {
    wb <- NULL
    return(FALSE)
    }
  # Save the file
  res <- tryCatch(saveWorkbook(wb), error = function(x) x)
  wb <- NULL
  return(TRUE)
  }

.writeToExcel.WriteXLS <- function(data, file, sheet, 
                                   header, rowheader, 
                                   envir, ...) {
  require(WriteXLS)
  WriteXLS(x = data,
           ExcelFileName = file,
           SheetNames = sheet,
           row.names = rowheader,
           col.names = header,
           envir = envir,
           ...)
  }

