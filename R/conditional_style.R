#' conditional_style
#'
#' object for conditional formatting
#' @import openxlsx
#' @include xlsx_builder_class.R
#' @export conditional_style
conditional_style<-setClass(
    "conditional_style",
    contains='xlsx_builder',
    slots=c(style='Style',
        cols='numeric',
        rows='numeric',
        rule='character',
        type='character')
)
