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
    ),
    validity=function(object) {
        A=all(rownames(object$row_header$data)==rownames(object$body$data))
        B=all(colnames(object$col_header$data)==colnames(object$body$data))
        
        valid=TRUE
        if (!A) {
            valid='Rownames of row header must match the rownames of the body.'
        }
        if (!B) {
            valid='Colnames of col header must match the colnames of the body.'
        }
        
        
        return(valid)
    }
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
    definition=function(S,...) {
        
        S=c(list(S),list(...))
        S=row_merge(S)
        return(S)
        
    })


row_merge=function(S) {
    # number of tables in sheet
    nt=length(S)
    
    # get unique row names over all tables
    u=rownames(S[[1]]$body$data)
    ru=rownames(S[[1]]$row_header$data)
    for (m in 2:nt) {
        u=unique(c(u,rownames(S[[m]]$body$data)))
        ru=unique(c(ru,rownames(S[[m]]$row_header$data)))
    }
    
    idx=rep(NA,length(u))
    
    # for each table, add any rows that are missing
    for (m in 1:nt) {
        
        # body
        M=matrix(NA,nrow=length(u),ncol=ncol(S[[m]]$body$data))
        M=as.data.frame(M)
        # row header
        r=matrix(NA,nrow=length(u),ncol=ncol(S[[m]]$row_header$data))
        r=as.data.frame(r)
        
        rownames(M)=u
        colnames(M)=colnames(S[[m]]$body$data)
        
        rownames(r)=ru
        colnames(r)=colnames(S[[m]]$row_header$data)
        
        # remove the names already present
        IN=u %in% rownames(S[[m]]$body$data)
        M=M[!IN,,drop=FALSE]
        r=r[!IN,,drop=FALSE]
        # bind the two tables
        M=rbind(S[[m]]$body$data,M)
        r=rbind(S[[m]]$row_header$data,r)
        # back into original order
        M=M[u,,drop=FALSE]
        r=r[ru,,drop=FALSE]
        # add back to table
        S[[m]]$body$data=M
        S[[m]]$row_header$data=r
        
        idx[IN]=min(idx[IN],m,na.rm=TRUE)
    }
    
    OUT=S[[1]]
    R=matrix(NA,nrow=length(ru),ncol=ncol(OUT$row_header$data))
    R=as.data.frame(R)
    rownames(R)=ru
    colnames(R)=colnames(OUT$row_header$data)
    
    for (k in 1:length(S)) {
        
        if (k>1) {
            OUT$body$data=cbind(OUT$body$data,S[[k]]$body$data,stringsAsFactors = FALSE)
            OUT$col_header$data=cbind(OUT$col_header$data,S[[k]]$col_header$data,stringsAsFactors = FALSE)
        }
        
        IN=rownames(S[[k]]$row_header$data) %in% ru[idx==k]
        R[idx==k,]=S[[k]]$row_header$data[IN,]
    }
    OUT$row_header$data=R
    
    return(OUT)
}



#' @export
setMethod(f='cmerge',signature=c('xlsx_table'),
    definition=function(S,...) {
        S=c(list(S),list(...))
        S=col_merge(S)
        return(S)
    })

col_merge=function(S) {
    # number of tables in sheet
    nt=length(S)
    
    # get unique row names over all tables
    u=colnames(S[[1]]$body$data)
    ru=colnames(S[[1]]$col_header$data)
    for (m in 2:nt) {
        u=unique(c(u,colnames(S[[m]]$body$data)))
        ru=unique(c(ru,colnames(S[[m]]$col_header$data)))
    }
    
    idx=rep(NA,length(u))
    
    # for each table, add any rows that are missing
    for (m in 1:nt) {
        
        # body
        M=matrix(NA,ncol=length(u),nrow=nrow(S[[m]]$body$data))
        M=as.data.frame(M)
        # row header
        r=matrix(NA,ncol=length(u),nrow=nrow(S[[m]]$col_header$data))
        r=as.data.frame(r)
        
        colnames(M)=u
        rownames(M)=rownames(S[[m]]$body$data)
        
        colnames(r)=ru
        rownames(r)=rownames(S[[m]]$col_header$data)
        
        # remove the names already present
        IN=u %in% colnames(S[[m]]$body$data)
        M=M[,!IN,drop=FALSE]
        r=r[,!IN,drop=FALSE]
        # bind the two tables
        M=cbind(S[[m]]$body$data,M)
        r=cbind(S[[m]]$col_header$data,r)
        # back into original order
        M=M[,u,drop=FALSE]
        r=r[,ru,drop=FALSE]
        # add back to table
        S[[m]]$body$data=M
        S[[m]]$col_header$data=r
        
        idx[IN]=min(idx[IN],m,na.rm=TRUE)
    }
    
    OUT=S[[1]]
    R=matrix(NA,ncol=length(ru),nrow=nrow(OUT$col_header$data))
    R=as.data.frame(R)
    colnames(R)=ru
    rownames(R)=rownames(OUT$col_header$data)
    
    for (k in 1:length(S)) {
        
        if (k>1) {
            OUT$body$data=rbind(OUT$body$data,S[[k]]$body$data,stringsAsFactors = FALSE)
            OUT$row_header$data=rbind(OUT$row_header$data,S[[k]]$row_header$data,stringsAsFactors = FALSE)
        }
        
        IN=colnames(S[[k]]$col_header$data) %in% ru[idx==k]
        R[,idx==k]=S[[k]]$col_header$data[,IN]
    }
    OUT$col_header$data=R
    
    return(OUT)
}

