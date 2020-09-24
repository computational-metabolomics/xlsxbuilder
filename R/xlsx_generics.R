#' @export
setGeneric("xlsx_write",function(WB,fn)standardGeneric("xlsx_write"))

#' @export
setGeneric("headers_from_body",function(x,mode)standardGeneric("headers_from_body"))

#' @export
setGeneric("as.xlsx_table",function(x)standardGeneric("as.xlsx_table"))

#' @export
setGeneric("body_style<-",function(x,value)standardGeneric("body_style<-"))

#' @export
setGeneric("body_border<-",function(x,value)standardGeneric("body_border<-"))

#' @export
setGeneric("body_border_style<-",function(x,value)standardGeneric("body_border_style<-"))

#' @export
setGeneric("row_style<-",function(x,value)standardGeneric("row_style<-"))

#' @export
setGeneric("row_border<-",function(x,value)standardGeneric("row_border<-"))

#' @export
setGeneric("row_border_style<-",function(x,value)standardGeneric("row_border_style<-"))

#' @export
setGeneric("col_style<-",function(x,value)standardGeneric("col_style<-"))

#' @export
setGeneric("col_border<-",function(x,value)standardGeneric("col_border<-"))

#' @export
setGeneric("col_border_style<-",function(x,value)standardGeneric("col_border_style<-"))

#' @export
setGeneric("merge_tables",function(S,mode,warn=TRUE)standardGeneric("merge_tables"))

#' @export
setGeneric("rmerge",function(S,...,all_equal=FALSE)standardGeneric("rmerge"))

#' @export
setGeneric("cmerge",function(S,...,all_equal=FALSE)standardGeneric("cmerge"))
