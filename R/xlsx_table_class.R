#' xlsx_table
#'
#' a table has (formatted) row and column names
#' @import openxlsx
#' @include xlsx_block_class.R xlsx_generics.R
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

#' @export
setMethod(f='rmerge',signature=c('xlsx_table'),
    definition=function(S,...,all_equal=FALSE) {
        
        S=c(list(S),list(...))
        if (length(S)>1) {
            S=row_merge(S,all_equal)
        }
        return(S)
        
    })


row_merge=function(S,all_equal=FALSE) {
    # S is a list of tables
    
    # number of tables in sheet
    nt=length(S)
    
    # new table
    OUT=S[[1]] # copy to get some formatting
    
    # merge body data
    M = S[[1]]$body$data
    colnames(M)=interaction(colnames(M),'1',sep='_')
    for (k in 2:nt) {
        N = S[[k]]$body$data
        colnames(N)=interaction(colnames(N),k,sep='_')
        M = merge(M,N,by='row.names',sort=FALSE,all=TRUE)
        rownames(M)=M$Row.names
        M$Row.names=NULL
    }
    OUT$body$data=M
    
    # row data
    OUT$row_header$data=data.frame('1'=rownames(M),row.names = rownames(M))
    
    # merge col data
    M = S[[1]]$col_header$data
    colnames(M)=interaction(colnames(M),'1',sep='_')
    for (k in 2:nt) {
        N = S[[k]]$col_header$data
        colnames(N)=interaction(colnames(N),k,sep='_')
        M = merge(M,N,by='row.names',sort=FALSE,all=TRUE)
        rownames(M)=M$Row.names
        M$Row.names=NULL
    }
    OUT$col_header$data=M

    
    
    # if row data is identical across tables then keep it
    M = S[[1]]$row_header$data
    all_equal = TRUE
    for (k in 2:nt) {
        
        N=S[[k]]$row_header$data
        
        if (!all(dim(M)==dim(N))) {
            all_equal=FALSE
            break
        }
        if (!all(M==N)) {
            all_equal=FALSE
            break
        }
    }
    # all equal so do replacement
    if (all_equal) {
        R=S[[1]]$row_header
        # sort the body int the same order as the header
        OUT$body$data=OUT$body$data[rownames(R$data),]
        OUT$row_header=R
    }
    return(OUT)
}



#' @export
setMethod(f='cmerge',signature=c('xlsx_table'),
    definition=function(S,...,all_equal=FALSE) {
        S=c(list(S),list(...))
        if (length(S)>1) {
            S=col_merge(S,all_equal)
        }
        return(S)
    })

col_merge=function(S,all_equal=FALSE) {
    
    # S is a list of tables
    
    # number of tables in sheet
    nt=length(S)
    
    # new table
    OUT=S[[1]]
    
    # merge body data
    M = t(S[[1]]$body$data)
    colnames(M)=interaction(colnames(M),1,sep='_')
    for (k in 2:nt) {
        N=t(S[[k]]$body$data)
        colnames(N)=interaction(colnames(N),k,sep='_')
        M = merge(M,N,by='row.names',sort=FALSE,all=TRUE)
        rownames(M)=M$Row.names
        M$Row.names=NULL
    }
    OUT$body$data=data.frame(t(M))
    
    # col data
    OUT$col_header$data=data.frame(matrix(colnames(OUT$body$data),nrow=1),row.names = '1')
    colnames(OUT$col_header$data)=colnames(OUT$body$data)
    
    # merge row data
    M = t(S[[1]]$row_header$data)
    colnames(M)=interaction(colnames(M),1,sep='_')
    for (k in 2:nt) {
        N=t(S[[k]]$row_header$data)
        colnames(N)=interaction(colnames(N),k,sep='_')
        M = merge(M,N,by='row.names',sort=FALSE,all=TRUE)
        rownames(M)=M$Row.names
        M$Row.names=NULL
    }
    OUT$row_header$data=data.frame(t(M))
    
    # if column data is identical across tables then keep it
    M = S[[1]]$col_header$data
    
    for (k in 2:nt) {
        
        N=S[[k]]$col_header$data
        
        if (!all(dim(M)==dim(N))) {
            all_equal=FALSE
            break
        }
        
        if (!all(M==N,na.rm = TRUE)) {
            all_equal=FALSE
            break
        }
    }
    # all equal so do replacement
    if (all_equal) {
        R=S[[1]]$col_header
        # sort the body into the same order as the header
        OUT$body$data=OUT$body$data[,colnames(R$data)]
        OUT$col_header=R
    }
    
    return(OUT)
}

