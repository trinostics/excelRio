Suppose you have a triangle in Excel that you would like to analyze in
R. In Excel, select the cells comprising the triangle *including the
labels identifying accident year and age* and copy to the clipboard.

<img src='smalltri.png' width="400px">

Back in R, load the **excelRio** package and run the function that
pastes content from the clipboard. Be sure to tell the function that
your data has a "header" for the column names (development ages) and a
"rowheader" for the row names (accident years).

    library(excelRio)

    ## Loading required package: tools

    tri <- pasteFromExcel(header = TRUE, rowheader = TRUE)

    ## Warning in read.table(file = file, header = header, sep = sep, quote =
    ## quote, : incomplete final line found by readTableHeader on 'clipboard'

    tri

    ## data frame with 0 columns and 0 rows

**excelRio** has special handling to strip out the dollar signs and
commas when appropriate. **excelRio**'s interpretation of "appropriate"
is, if after stripping out the dollar signs and commas the column of
results can be successfully converted to **numeric**, then that's what
will happen. Without special handling, this would not normally happen,
as below when the first two character elements lose numeric value.

    as.numeric(c("$1,000", "1,200", "800"))

    ## Warning: NAs introduced by coercion

    ## [1]  NA  NA 800

The argument that controls this special handling is
**convertFormattedNumbers** and can be turned off, as in the code below.
Note that I also set **stringsAsFactors = FALSE** because my
installation of **R** has the default value TRUE for that global option.

    tri <- pasteFromExcel(header = TRUE, rowheader = TRUE, 
                          convertFormattedNumbers = FALSE, 
                          stringsAsFactors = FALSE)

    ## Warning in read.table(file = file, header = header, sep = sep, quote =
    ## quote, : incomplete final line found by readTableHeader on 'clipboard'

    tri

    ## data frame with 0 columns and 0 rows

    mode(tri)

    ## [1] "list"

    class(tri)

    ## [1] "data.frame"

The *class* of the above raises a subtle point about the value of the
result from **pasteFromExcel**.

Data in Excel is often comprised of columns of data of many different
types: numeric, character, and dates, to name a few. By default,
therefore, **pasteFromExcel** will create a data.frame, an **R** object
that allows differently typed columns. But when all the columns of data
are "compatible," **pasteFromExcel** will return a matrix object of the
common "type" (or "mode") by default, if possible, under the assumption
that a matrix object would be "intuitively expected." The argument that
controls that option is **simplify**. By calling the function with
**simplify = FALSE**, a data.frame will always be returned.

For more information about **pasteFromExcel** and its various optional
arguments, see

    ?pasteFromExcel

Finally, Dear Reader, I leave you with a warning and a request for
assistance! :)

### *Warning!*

Formatted numbers in Excel are stripped of their "adornments" via a
[regular expression](http://en.wikipedia.org/wiki/Regular_expression) I
found on the internet
[here](http://stackoverflow.com/questions/354044/what-is-the-best-u-s-currency-regex).
**This expression will not work outside the U.S.**

I repeat:

***Warning! excelRio's special currency handling will only work with
dollar signs, with periods as decimal separator, and with commas
separating triples.*** I apologize. :(

(For the curious, **excelRio** stores its current regex pattern for
currency in the variable **currencypattern** on line 38 in the file
**excelRio.r** in the R folder of the package on github. )

### *Request for Assistance!*

If anyone is interested in suggesting a currency pattern for their
location I would really appreciate the collaboration --- woohoo!! You
can reach me at danielmarkmurphy at gmail. Thank you sincerely in
advance.

Dan

* * * * *

created (mostly) with RMarkdown in RStudio
