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

#' @export
setMethod(f='merge_tables',signature=c('xlsx_sheet','character'),
  definition=function(S,mode,warn=TRUE) {
    
    # check for correct mode input
    if (!(mode %in% c('cols','rows'))) {
      stop('incorrect mode specified. Can only be "cols" or "rows"')
    }
    
    # number of tables in sheet
    nt=length(S$tables)
    
    if (mode =='rows') {
      # get unique row names over all tables
      u=rownames(S$tables[[1]]$body$data)
      for (m in 2:nt) {
        u=unique(c(u,rownames(S$tables[[m]]$body$data)))
      }
      # for each table, add any rows that are missing
      for (m in 1:nt) {
        M=matrix(NA,nrow=length(u),ncol=ncol(S$tables[[m]]$body$data))
        M=as.data.frame(M)
        rownames(M)=u
        colnames(M)=colnames(S$tables[[m]]$body$data)
        # remove the names already present
        IN=rownames(S$tables[[m]]$body$data) %in% u
        M=M[-IN,,drop=FALSE]
        # bind the two tables
        M=rbind(S$tables[[m]]$body$data,M)
        # arrage according to u
        M=M[u,,drop=FALSE]
        # add back to table
        S$tables[[m]]$body$data=M
      }
      # replace row_header for first table, and remove for the other tables
      r=as.data.frame(matrix(0,nrow=length(u),ncol=1))
      r[,1]=u
      R=S$tables[[1]]$row_header
      R$data=r
      
      if (warn) {
        if (ncol(S$tables[[m]]$row_header$data)>0) {
          warning('Row headers will be replaced by merged row names based on the row names of "body" data frames')
        }
      }
      S$tables[[1]]$row_header=R
      
      for (m in 2:nt) {
        S$tables[[m]]$row_header=xlsx_block()
      }
    }
    
    if (mode=='cols') {
      stop('"cols" merge not implemented yet')
    }
    
    return(S)
  })
