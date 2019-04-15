#' xlsx_sheet
#'
#' a sheet combines multiple tables.
#' @import openxlsx
#' @include xlsx_table_class.R
#' @export xlsx_sheet
xlsx_sheet<-setClass(
    "xlsx_sheet",
    contains='xlsx_builder',
    slots=c(tables='list',
        merge='logical',
        name='character',padding='numeric'),
    prototype=list(merge=FALSE,padding=0)
)

#' @export
setMethod("+",
    signature(e1='xlsx_table',e2='xlsx_table'),
    definition=function(e1,e2) {
        OUT=xlsx_sheet(tables=c(e1,e2))
        return(OUT)
    }
)

#' @export
setMethod("+",
    signature(e1='xlsx_sheet',e2='xlsx_table'),
    definition=function(e1,e2) {
        e1$tables=c(e1$tables,e2)
        return(e1)
    }
)

#' @export
setMethod("+",
    signature(e1='xlsx_table',e2='xlsx_sheet'),
    definition=function(e1,e2) {
        e2$tables=c(e1,e2$tables)
        return(e2)
    }
)
