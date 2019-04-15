#' xlsx_table
#'
#' a table has (formatted) row and column names
#' @import openxlsx
#' @include xlsx_block_class.R
#' @export xlsx_table
xlsx_table<-setClass(
    "xlsx_table",
    contains='xlsx_builder',
    slots=c(body='xlsx_block',
            col_header='xlsx_block',
            row_header='xlsx_block'
        )
)

