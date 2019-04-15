#' xlsx_block
#'
#' data that goes somewhere in an excel sheet
#' @import openxlsx
#' @include xlsx_builder_class.R conditional_style.R
#' @export xlsx_block
xlsx_block<-setClass(
    "xlsx_block",
    contains='xlsx_builder',
    slots=c(data='data.frame',
        style='Style',
        border='character',
        border_style='character',
        conditional='conditional_style',
        keep_na='logical'),
    prototype = list(border='none',keep_na=FALSE)
)

