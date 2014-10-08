\name{copyToExcel}
\alias{copyToExcel}
\title{Copy data from R to Excel via the clipboard}
\description{
A fast way to transfer small amounts of data from \R to Excel.
}
\usage{
copyToExcel(x, rowheader = TRUE, header = TRUE, na = "", \dots)
}
\arguments{
    \item{x}{name of \R object to write out}
    \item{rowheader}{did you want to write out the row header(s)/name(s) too?}
    \item{header}{did you want to write out the column header(s) too?}
    \item{na}{the string to use for missing values in the data}
    \item{\dots}{arguments to be passed to other methods}
    }
\details{
Call this function with your data as the x argument and its values
will be written to the clipboard.
In Excel, 
"paste" into the location whose upper left corner is your cursor's position.
Be careful where you paste because 
\emph{contents will be overwritten and there is no "undo"!}

This is a wrapper for R's write.table function that standardizes 
the argument list throughout the excelRio package.
It also calls write.table with the appropriate value for \code{file}
depending on the two supported operating systems: Windows and Debian.

The default behavior for argument \code{na} changes the default behavior
of write.table so that NA values show as white space in Excel
rather than as "NA".
}
\value{
NULL
}
\section{Warning}{
Contents will be overwritten and there is no "undo"!
}
\author{
dmm
}
\seealso{
pasteFromExcel, writeToExcel, writeToCsv
}
\examples{
loss <- 100 * (1:5)
names(loss) <- LETTERS[1:5]
### NOT RUN
# copyToExcel(loss) # Saves the data to the clipboard.
# In Excel "paste" the data where you want it.
}
\keyword{ csv }
