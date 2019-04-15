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
        name='character')
)

