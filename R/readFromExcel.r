# readFromExcel.r
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
    
readFromExcel <- function(file = choose.files(), sheet = 1, 
  stringsAsFactors = default.stringsAsFactors(),
  simplify = TRUE, drop = TRUE, 
  na.strings=c("NA","#DIV/0!"), 
  zero.strings = "-", 
  convertFormattedNumbers = TRUE, ..., 
  header = TRUE, rowheader = FALSE, 
  pkg = c("XLConnect", "RODBC", "xlsx")) {
  
  # Read data from a worksheet in an Excel file.
  # Oct. 15, 2012

  if (missing(file)) {
    file <- choose.files()
    if (length(file) == 0) # hit Cancel
      return(file)
    }
#  require(tools, quietly = TRUE)
  ext <- tolower(tools::file_ext(file))
  if (!(ext %in% c("xls", "xlsx", "xlsm", ""))) {
    cat(file, "' is not a supported Excel file\n", sep = "")
    return(NULL)
    }
  # 7/19/2014 If file.ext is blank, try xlsx, xls,and xlsm in that order
  fyle <- file
  if (ext == "") file <- paste(fyle, "xlsx", sep = ".")
  if (!file.exists(file)) file <- paste(fyle, "xls", sep = ".")
  if (!file.exists(file)) file <- paste(fyle, "xlsm", sep = ".")
  if (!file.exists(file)) stop("File '", fyle, "' does not exist.", sep = "")

  pkg <- match.arg(pkg)
  if (pkg == "gdata") stop ("gdata currently unsupported due to bug")  

  # If XLConnect is an installed package, we'll try that first.
  if (pkg == "XLConnect") {
    if ("XLConnect" %in% installed.packages()[,1]) {
#      res <- suppressPackageStartupMessages(require(XLConnect, warn.conflicts = FALSE, quietly = TRUE))

      res <- wb <- tryCatch(XLConnect::loadWorkbook(file, create = FALSE), error = function(e) e)
      if (class(res)[1L] != "workbook") {
        warning("Problem '", res$message, "' trying to load workbook '", file, "' with XLConnect package. Can happen when also open.\n",
                "  If out of memory, try pkg = 'RODBC'\n")
        return(res)
        }
      # Let's see if the sheet exists
      shts <- XLConnect::getSheets(wb)
      if (is.numeric(sheet)) {
        if (sheet > length(shts)) {
          cat("You asked for sheet number ", sheet, " but there are only ", length(shts), " sheets in file '", file, "'\n", sep = "")
          return(NULL)
          }
        }
      else
      if (!(sheet %in% shts)) {
        cat("Read failed. Sheet '", sheet, "' does not exist in file '", file, "'\n", sep = "")
        return(NULL)
        }
      # rownames doesn't seem to work as expected in the next line, so commented out
      #x <- tryCatch(readWorksheet(object = wb, sheet = sheet, header = header, rownames = if (rowheader) 1 else NULL), error = function(x) x)
      x <- tryCatch(XLConnect::readWorksheet(object = wb, sheet = sheet, header = header, rownames = NULL, ...), error = function(x) x)
      if ("error" %in% class(x)) stop("Problem reading '", file, "', sheet '", sheet, "' with XLConnect package.")
      if (rowheader) {
        row.names(x) <- x[[1]]
        x[1] <- NULL
        }
      if (convertFormattedNumbers | stringsAsFactors) for (i in which(sapply(x, is.character))) {
        if (convertFormattedNumbers) {
          is.n <- is.numeric.string(x[[i]], zero.strings = zero.strings)
          if (is.n) {
            x[[i]] <- attr(is.n, "value")
            next
            }
          x[[i]] <- as.factor(x[[i]])
          }
        }
      
      # 8/2/2014 Implement stringsAsFastors = FALSE
      if (!stringsAsFactors) {
        classes <- sapply(x, function(z) class(z)[1L])
        ifac <- which(classes == "factor")
        for (i in ifac) x[[i]] <- as.character(x[[i]])
        }
      # 8/4/2014 Implement identify date string
      for (i in seq_along(x)) {
        if (w <- is.date.string(x[[i]])) x[[i]] <- attr(w, "value")
        }
      attr(x, "basename") <- basename(file)
      attr(x, "dirname") <- dirname(file)
      attr(x, "fullname") <- file.path(attr(x, "dirname"), attr(x, "basename"))
      attr(x, "file_ext") <- tools::file_ext(file)
      if (simplify) x <- .simplifyDataFrame(x, drop = drop)
      return(x)
      }
    else stop("XLConnect not an installed package.")
    }
    
  # If gdata is an installed package, we'll try that next.
  # Oct 22, 2012: bug reading negative numbers formatted Accounting in Excel.
  #   Skip until fixed
  if (FALSE) {
#  if (pkg == "gdata") {
    if ("gdata" %in% installed.packages()[,1]) {
#      res <- suppressPackageStartupMessages(require(gdata, warn.conflicts = FALSE, quietly = TRUE))
  
      # Let's see if the sheet exists
      shts <- sheetNames(file)
      if (is.numeric(sheet)) {
        if (sheet > length(shts)) {
          cat("You asked for sheet number ", sheet, " but there are only ", length(shts), " sheets in file '", file, "'\n", sep = "")
          return(NULL)
          }
        }
      else
      if (!(sheet %in% shts)) {
        cat("Read failed. Sheet '", sheet, "' does not exist in file '", file, "'\n", sep = "")
        return(NULL)
        }
      x <- tryCatch(read.xls(xls = file, sheet = sheet, stringsAsFactors = stringsAsFactors, header = header, ...), error = function(x) x)
      if ("error" %in% class(x)) stop("Problem reading '", file, "', sheet'", sheet, "' with gdata package.")
      if (rowheader) {
        row.names(x) <- x[[1]]
        x[1] <- NULL
        }
      attr(x, "basename") <- basename(file)
      attr(x, "dirname") <- dirname(file)
      attr(x, "fullname") <- file.path(attr(x, "dirname"), attr(x, "basename"))
      attr(x, "file_ext") <- tools::file_ext(file)
      if (simplify) x <- .simplifyDataFrame(x, drop = drop)

      return(x)
      }
    else stop("gdata not an installed package.")
    }
    
  # If RODBC is an installed package, we'll try that next.
  #   Works best if columns names are in the first row of the sheet.
  if (pkg == "RODBC") {
    if ("RODBC" %in% installed.packages()[,1]) {
      # if (R.Version()$arch != "i386") stop("Regretably, RODBC only works in 32 bit R.")
      # Actually, it's more complicated. Depends on installed version of excel.
#      res <- suppressPackageStartupMessages(require(RODBC, warn.conflicts = FALSE, quietly = TRUE))
      
      channel <- tryCatch(if (ext == "xls") RODBC::odbcConnectExcel(file) else RODBC::odbcConnectExcel2007(file), error = function(x) x)
      if (class(channel) != "RODBC") stop("Problem trying to connect to workbook '", file, 
        "' with RODBC package.\n\nCan happen when\n", 
        "1. the Excel file is currently open\n", 
        "2. the 32/64 bit version of R does not match the 32/64 bit version of Excel\n", 
        "   your architecture is '", R.Version()$arch, "'\n\n")
  
      # Define getSheets and existsSheet functions for RODBC
      getSheets <- function() {
        xltbls <- RODBC::sqlTables(channel)
        shts <- xltbls$TABLE_NAME[grep("\\$", xltbls$TABLE_NAME)]
        if (length(shts) == 0) stop("No sheets for RODBC to read!")
        # Get rid of the dollar sign at the end of the name that signifies that this excel "table" is a sheet
        shts <- gsub('\\$', "", shts)
        # Get rid of single quotes at the beginning and the end, if any
        shts <- gsub("'$", "", gsub("^'", "", shts))
        }
      existsSheet <- function(sheet) sheet %in% getSheets()
      # Let's see if the sheet exists
      if (is.numeric(sheet)) {
        shts <- getSheets()
        if (sheet > length(shts)) {
          cat("You asked for sheet number ", sheet, " but there are only ", length(shts), " sheets in file '", file, "'\n", sep = "")
          RODBC::odbcClose(channel)
          return(NULL)
          }
        # Keep fingers crossed that the sheet # user asks for is the one RODBC returns
        sheet <- shts[sheet] # store the name of the sheet instead of the sheet #
        }
      else
      if (!existsSheet(sheet)) {
        cat("Read failed. Sheet '", sheet, "' does not exist in file '", file, "'\n", sep = "")
        RODBC::odbcClose(channel)
        return(NULL)
        }
  
      # As long as colnames = FALSE, sqlFetch seems to do the right thing
      #   regarding interpreting first row of data as column names if that 
      #   makes sense or, if not, making its own column names
      # 'rownames' argument doesn't seem to work as expected in the next line, so commented out
      #x <- tryCatch(sqlFetch(channel, sheet, stringsAsFactors = stringsAsFactors, na.strings = na.strings, ..., colnames = FALSE, rownames = rowheader), error = function(e) e)
      x <- tryCatch(RODBC::sqlFetch(channel, sheet, stringsAsFactors = stringsAsFactors, na.strings = na.strings, ..., colnames = FALSE, rownames = FALSE), error = function(e) e)
      RODBC::odbcClose(channel)
      if ("error" %in% class(x)) stop("Problem reading '", file, "', sheet '", sheet, "' with RODBC package. File currently open in Excel?")
      if (rowheader) {
        row.names(x) <- x[[1]]
        x[1] <- NULL
        }
      attr(x, "basename") <- basename(file)
      attr(x, "dirname") <- dirname(file)
      attr(x, "fullname") <- file.path(attr(x, "dirname"), attr(x, "basename"))
      attr(x, "file_ext") <- tools::file_ext(file)
      if (simplify) x <- .simplifyDataFrame(x, drop = drop)
      return(x)
      }
    else stop("RODBC not an installed package.")
    }

  # last resort is the xlsx package   
  if (all(c("rJava", "xlsx") %in% installed.packages()[,1])) {
#    require(rJava, warn.conflicts = FALSE, quietly = TRUE)
#    require(xlsx, warn.conflicts = FALSE, quietly = TRUE)
    res <- wb <- tryCatch(xlsx::loadWorkbook(file), error = function(x) x)
    if (class(res) != "jobjRef") stop("Problem trying to load workbook '", file, "' with gdata package.")

    # Define existsSheet function for xlsx
    existsSheet <- function(sheet) sheet %in% names(getSheets(wb))

    # Let's see if the sheet exists
    if (is.numeric(sheet)) {
      shts <- names(xlsx::getSheets(wb))
      if (sheet > length(shts)) {
        cat("You asked for sheet number ", sheet, " but there are only ", length(shts), " sheets in file '", file, "'\n", sep = "")
        return(NULL)
        }
      sheet <- shts[sheet] # store the name of the sheet instead of the sheet #
      }
    else
    if (!xlsx::existsSheet(sheet)) {
      cat("Read failed. Sheet '", sheet, "' does not exist in file '", file, "'\n", sep = "")
      return(NULL)
      }
    wb <- NULL # destroy utility variable
    x <- tryCatch(xlsx::read.xlsx(file = file, sheetName = sheet, rowIndex=NULL,
          colIndex=NULL, as.data.frame=TRUE, header=header, colClasses=NA,
          keepFormulas=FALSE, encoding="unknown", stringsAsFactors = stringsAsFactors, ...), 
       error = function(x) x)
       
    if ("error" %in% class(x)) stop("Problem reading '", file, "', sheet'", sheet, "' with xlsx package.")
    if (rowheader) {
      row.names(x) <- x[[1]]
      x[1] <- NULL
      }
    attr(x, "basename") <- basename(file)
    attr(x, "dirname") <- dirname(file)
    attr(x, "fullname") <- file.path(attr(x, "dirname"), attr(x, "basename"))
    attr(x, "file_ext") <- tools::file_ext(file)
    if (simplify) x <- .simplifyDataFrame(x, drop = drop)
    return(x)
    }
  stop("Packages gdata, RODBC, and xlxs unavailable for reading Excel files. Try csv or clipboard exchange.")
  }

