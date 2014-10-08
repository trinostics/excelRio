# excelRio 

Excel - R Input - Output

excelRio is an R package providing rudimentary input/output capability
between Microsoft Excel and R via:

- the clipboard
- csv files
- directly to workbooks

In contraposition to RExcel, excelRio assumes the user "lives" within the R environment
and sometimes uses Excel as a source of data or to store results.

## Installation

Currently this package is under development and only exists on github.

To install the current development version from github you need the [devtools package](http://cran.r-project.org/web/packages/devtools/index.html) and the other packages on which excelRio depends:

```s
install.packages(c("tools", "mondate"))
```

Those two packages are sufficient to exchange data via the clipboard and csv files.
To exchange data directly with workbooks it is recommended to also install the XLConnect package and it dependencies.
Note that XLConnect uses java and the rJava package, which can sometimes be troublesome to manage. Please consult your IT department.

To install excelRio run:
```s
library(devtools)
install_github("trinostics/excelRio")
```

## Usage

```s
library(excelRio)
?excelRio
```

Refer to the help files. Examples are somewhat limited at this time. 