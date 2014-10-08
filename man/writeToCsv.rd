\name{writeToCsv}
\alias{writeToCsv}
\title{Write data from an \R object to a csv file}
\description{
Use this function to 
the values of an \R variable to a csv file.
}
\usage{
  writeToCsv(data, file = deparse(substitute(data)), 
            header, rowheader = TRUE, 
            overwrite = FALSE, na = "", \dots)
  }
\arguments{
    \item{data}{name of \R object to write out}
    \item{file}{file to write to; default is the name of the \R object}
    \item{header}{do you want to write out the column header(s)/names(s) too?}
    \item{rowheader}{do you want to write out the row header(s)/name(s) too?}
    \item{overwrite}{
      If TRUE, file will always be overwritten.
      If FALSE, file will never be overwritten.
      If "version", the first version of the filename of the 
      form 'file-xx.csv' that doesn't yet exist will be created,
      up to a maximum of xx=99.
      }
    \item{na}{the string to use for missing values in the data}
    \item{\dots}{arguments to be passed to other methods}
    }
\details{
Essentially a wrapper for R's write.csv function that standardizes 
the argument list throughout the excelRio package, 
does some error checking,
and enables a simple "version-ing" for different iterations
of the same \code{data} variable.
}
\value{
The name with path of the file written.
}
\author{
dmm
}
\seealso{
readFromCsv, pasteFromExcel, copyToExcel
}
\examples{
loss <- 100 * (1:5)
names(loss) <- LETTERS[1:5]
### NOT RUN
# writeToCsv(loss) # Saves the data in 'dataout.csv'
}
\keyword{ csv }
