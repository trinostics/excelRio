\name{writeToExcel}
\alias{writeToExcel}
\title{Write data from an \R object to an Excel sheet}
\description{
In Excel, 
save data to a csv file.
Uuse this function to 
read that data into an \R variable.
}
\usage{
  writeToExcel(data, file = deparse(substitute(data)), 
               sheet = "Sheet1", 
               startRow = 1, startCol = 1, 
               header = TRUE, rowheader = TRUE, 
               overwrite = TRUE,
               pkg = "XLConnect",
               \dots)
  }
\arguments{
    \item{data}{name of \R object to write out}
    \item{file}{file to write to; default is the name of the \R object}
    \item{sheet}{the sheet name or number}
    \item{startRow}{the top row in which to begin storing the output}
    \item{startCol}{the leftmost column in which to begin storing the output}
    \item{header}{do you want to write out the column header(s)/names(s) too?}
    \item{rowheader}{do you want to write out the row header(s)/name(s) too?}
    \item{overwrite}{
      If TRUE, file will always be overwritten.
      If FALSE, file will never be overwritten.
      If "version", the first version of the filename of the 
      form 'file-xx.csv' that doesn't yet exist will be created,
      up to a maximum of xx=99.
      }
    \item{pkg}{the name of a supported package, 
      currently "XLConnect", the default, or "WriteXLS". See Details.}
    \item{\dots}{arguments to be passed to other methods}
    }
\details{
Essentially a wrapper for R packages XLConnect or WriteXLS.

XLConnect is more versatile:
can write new content to an existing file;
can write new content to an existing sheet;
can write content to a specific area of a sheet.
In that sense, XLConnect is useful as "report writer" of output
to an Excel file that will eventually be printed, possibly after
some manipulation.
On the downside, 
it requires java, which can be difficult to install and support
on a machine without the proper authorities.
Java constraints also limit the size of the objects that can be written.

WriteXLS does not use java so can write out larger objects.
It is more limited in the sense that it completely replaces
an existing file, so new sheets cannot be added to a workbook.
Also, objects are written beginning in row 1, column 1, 
so it is not possible to write to a specific region in a sheet.
In that sense, WriteXLS is better used to dump large amounts of data to
Excel, but arguably writeToCsv is a more capable tool.

\code{overwrite = "version"} is a convenient way to incrementally store results 
from different versions of a model.
  
}
\value{
The name with path of the file written.
}
\author{
dmm
}
\seealso{
readFromExcel, writeToCsv, copyToExcel
}
\examples{
loss <- 100 * (1:5)
names(loss) <- LETTERS[1:5]
### NOT RUN
# writeToExcel(loss) # Saves the data in 'dataout.csv'
}
\keyword{ excel }
