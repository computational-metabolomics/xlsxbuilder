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
        row_header='xlsx_block',
        name='character'
    )
)

#' @export headers_from_body
setMethod(f="headers_from_body",
    signature(x='xlsx_table'),
    definition=function(x,mode='both') {
        if (mode %in% c('both','col')) {
            h=as.data.frame(matrix(NA,1,ncol(x$body$data)))
            h[1,]=colnames(x$body$data)
            h=xlsx_block(data=h,style=createStyle(textDecoration = 'bold'),border='surrounding',border_style = 'thin')
            x$col_header=h
        }
        if (mode %in% c('both','row')) {
            r=as.data.frame(matrix(NA,nrow(x$body$data),1))
            r[,1]=rownames(x$body$data)
            r=xlsx_block(data=r,style=createStyle(textDecoration = 'bold'),border='surrounding',border_style = 'thin')
            x$row_header=r
        }
        return(x)
    }
)

#' @export
setMethod(f='as.xlsx_table',
    signature(x='data.frame'),
    definition = function (x){
        body=xlsx_block(data=x)
        out=xlsx_table(body=body)
        out=headers_from_body(out,'both')
        return(out)
    }
)

#' @export
setMethod(f='as.xlsx_table',
    signature(x='matrix'),
    definition = function (x){
        body=xlsx_block(data=as.data.frame(x))
        out=xlsx_table(body=body)
        out=headers_from_body(out,'both')
        return(out)
    }
)

#' @export
setMethod(f='body_style<-',
    signature(x='xlsx_table'),
    definition = function (x,value){
        x$body$style=value
        return(x)
    }
)

#' @export
setMethod(f='body_border<-',
    signature(x='xlsx_table'),
    definition = function (x,value){
        x$body$border=value
        return(x)
    }
)

#' @export
setMethod(f='body_border_style<-',
    signature(x='xlsx_table'),
    definition = function (x,value){
        x$body$border_style=value
        return(x)
    }
)

#' @export
setMethod(f='row_style<-',
    signature(x='xlsx_table'),
    definition = function (x,value){
        x$row_header$style=value
        return(x)
    }
)

#' @export
setMethod(f='row_border<-',
    signature(x='xlsx_table'),
    definition = function (x,value){
        x$row_col$border=value
        return(x)
    }
)

#' @export
setMethod(f='row_border_style<-',
    signature(x='xlsx_table'),
    definition = function (x,value){
        x$row$border_style=value
        return(x)
    }
)

#' @export
setMethod(f='col_style<-',
    signature(x='xlsx_table'),
    definition = function (x,value){
        x$col_header$style=value
        return(x)
    }
)

#' @export
setMethod(f='col_border<-',
    signature(x='xlsx_table'),
    definition = function (x,value){
        x$col_header$border=value
        return(x)
    }
)

#' @export
setMethod(f='col_border_style<-',
    signature(x='xlsx_table'),
    definition = function (x,value){
        x$col_header$border_style=value
        return(x)
    }
)
