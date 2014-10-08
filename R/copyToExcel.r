# copyToExcel.r
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
    
copyToExcel <- function(x, rowheader = TRUE, header = TRUE, na = "", ...) {
  y <- as.data.frame(x)
  if (is.vector(x)) {
    names(y) <- deparse(substitute(x))
    if (missing(rowheader)) rowheader <- !is.null(names(x))
    }
  switch(Sys.info()["sysname"],
    Windows = write.table(y, file = "clipboard", sep = "\t", row.names = rowheader, col.names = if (rowheader && header) NA else header, ...),
    Darwin = {
      clip <- pipe("pbcopy", "w")
      write.table(y, file = clip, sep = "\t", row.names = rowheader, col.names = if (rowheader && header) NA else header, na = na, ...)
      close(clip)
      },
    stop("unsupported OS")
    )
  }

