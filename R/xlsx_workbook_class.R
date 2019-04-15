#' xlsx_workbook
#'
#' a workbook combines multiple sheets.
#' @import openxlsx
#' @include xlsx_sheet_class.R
#' @export xlsx_workbook
xlsx_workbook<-setClass(
    "xlsx_workbook",
    contains='xlsx_builder',
    slots=c(sheets='list',overwrite='logical'),
    prototype=list(overwrite=FALSE)
)

#' @export
setMethod(f="xlsx_write",
    signature=c("xlsx_workbook",'character'),
    definition=function(WB,fn) {

        # prepare workbook
        wb=createWorkbook()

        # add sheets
        ns=length(WB$sheets)
        for (k in 1:ns) {
            col_count = 0 # current column number
            row_count = 0 # current row number
            # for each sheet
            nt=length(WB$sheets[[k]]$tables)
            # add the sheet
            addWorksheet(wb,WB$sheets[[k]]$name)
            # count the number of header rows
            head_count=numeric(nt)
            data_count=numeric(nt)
            for(m in 1:nt) {
                head_count[m]=nrow(WB$sheets[[k]]$tables[[m]]$col_header$data)
            }

            # pad the headers of smaller tables with spaces to make them the same size as the largest one
            for(m in 1:nt) {
                if (head_count[m]<max(head_count)) {
                    temp=matrix(NA,nrow=max(head_count[m]-head_count[m]),ncol=ncol(WB$sheets[[k]]$tables[[m]]$col_header$data))
                    temp=as.data.frame(temp)
                    colnames(temp)=colnames(WB$sheets[[k]]$tables[[m]]$col_header$data)
                    WB$sheets[[k]]$tables[[m]]$col_header$data=rbind(temp,WB$sheets[[k]]$tables[[m]]$col_header$data)
                }
            }

            # if merge = true, attempt to (row) merge the tables in a sheet
            # get unique row names across all tables
            if (WB$sheets[[k]]$merge) {
                u=rownames(WB$sheets[[k]]$tables[[1]]$body$data)
                for (m in 2:nt) {
                    u=unique(c(u,rownames(WB$sheets[[k]]$tables[[m]]$body$data)))
                }
                # for each table, add any rows that are missing
                for (m in 1:nt) {
                    M=matrix(NA,nrow=length(u),ncol=ncol(WB$sheets[[k]]$tables[[m]]$body$data))
                    M=as.data.frame(M)
                    rownames(M)=u
                    colnames(M)=colnames(WB$sheets[[k]]$tables[[m]]$body$data)
                    # remove the names already present
                    IN=rownames(WB$sheets[[k]]$tables[[m]]$body$data) %in% u
                    M=M[-IN,]
                    # bind the two tables
                    M=rbind(WB$sheets[[k]]$tables[[m]]$body$data,M)
                    # arrage according to u
                    M=M[u,]
                    # add back to table
                    WB$sheets[[k]]$tables[[m]]$body$data=M
                }
                # replace row_header for first table, and remove for the other tables
                R=as.data.frame(matrix(0,nrow=length(u),ncol=1))
                R[,1]=u
                R=xlsx_block(data=R,style=createStyle(textDecoration = 'bold',fontColour ='#0000ff'),border='columns',border_style='thin')

                if (ncol(WB$sheets[[k]]$tables[[m]]$row_header$data)>0) {
                    warning('Row headers will be replaced by merged row names based on the row names of "body" data frames')
                }
                WB$sheets[[k]]$tables[[1]]$row_header=R

                for (m in 2:nt) {
                    WB$sheets[[k]]$tables[[m]]$row_header=xlsx_block()
                }
            }

            for (m in 1:nt) {
                # for each table

                hheight=nrow(WB$sheets[[k]]$tables[[m]]$col_header$data)
                rwidth=ncol(WB$sheets[[k]]$tables[[m]]$row_header$data)
                rheight=nrow(WB$sheets[[k]]$tables[[m]]$row_header$data)
                hwidth=ncol(WB$sheets[[k]]$tables[[m]]$col_header$data)
                bheight=nrow(WB$sheets[[k]]$tables[[m]]$body$data)
                bwidth=ncol(WB$sheets[[k]]$tables[[m]]$body$data)

                # put the row header
                if (rwidth>0) {
                    writeData(wb=wb,sheet=WB$sheets[[k]]$name,
                        startRow = row_count+hheight+1,
                        startCol= col_count + 1,
                        x = WB$sheets[[k]]$tables[[m]]$row_header$data,
                        colNames = FALSE,borders = WB$sheets[[k]]$tables[[m]]$row_header$border,
                        borderStyle = WB$sheets[[k]]$tables[[m]]$row_header$border_style)
                    # add formatting
                    addStyle(wb = wb,sheet = WB$sheets[[k]]$name,
                        style = WB$sheets[[k]]$tables[[m]]$row_header$style,
                        rows=1:rheight+hheight+row_count,
                        cols=1:rwidth+col_count,
                        gridExpand = TRUE,stack=TRUE
                    )

                    # add conditional if set
                    C=WB$sheets[[k]]$tables[[m]]$row_header$conditional
                    if (length(C$rule)>0) {
                        conditionalFormatting(wb=wb,sheet = WB$sheets[[k]]$name,
                            rows=C$rows+hheight+row_count,
                            cols=C$cols+col_count,rule=C$rule,style=C$style,
                            type=C$type)
                    }
                }

                # put the column header
                if (hwidth>0) {
                    writeData(wb=wb,sheet=WB$sheets[[k]]$name,
                        startRow = row_count+1,
                        startCol=rwidth+col_count+1,
                        x = WB$sheets[[k]]$tables[[m]]$col_header$data,
                        colNames = FALSE,borders = WB$sheets[[k]]$tables[[m]]$col_header$border,
                        borderStyle = WB$sheets[[k]]$tables[[m]]$col_header$border_style)
                    # add formatting
                    addStyle(wb = wb,sheet = WB$sheets[[k]]$name,
                        style = WB$sheets[[k]]$tables[[m]]$col_header$style,
                        rows=1:hheight+row_count,
                        cols=1:hwidth+rwidth+col_count,
                        gridExpand = TRUE,stack=TRUE
                    )
                    # add conditional if set
                    C=WB$sheets[[k]]$tables[[m]]$col_header$conditional
                    if (length(C$rule)>0) {
                        conditionalFormatting(wb=wb,sheet = WB$sheets[[k]]$name,
                            rows=C$rows+row_count,
                            cols=C$cols+rwidth+col_count,style=C$style,
                            type=C$type,
                            rule=C$rule)
                    }
                }

                # put the data
                if (bwidth>0) {
                    writeData(wb=wb,sheet=WB$sheets[[k]]$name,
                        startRow = row_count+hheight+1,
                        startCol= col_count+rwidth+1,
                        colNames=FALSE,borders = WB$sheets[[k]]$tables[[m]]$body$border,
                        borderStyle = WB$sheets[[k]]$tables[[m]]$body$border_style,
                        x=WB$sheets[[k]]$tables[[m]]$body$data,keepNA = WB$sheets[[k]]$tables[[m]]$body$keep_na)
                    # add formatting
                    addStyle(wb = wb,sheet = WB$sheets[[k]]$name,
                        style = WB$sheets[[k]]$tables[[m]]$body$style,
                        rows=1:bheight+hheight+row_count,
                        cols=1:bwidth+rwidth+col_count,
                        gridExpand = TRUE,stack=TRUE
                    )
                    # add conditional if set
                    C=WB$sheets[[k]]$tables[[m]]$body$conditional
                    if (length(C$rule)>0) {
                        conditionalFormatting(wb=wb,sheet = WB$sheets[[k]]$name,
                            rows=C$rows+hheight+row_count,
                            cols=C$cols+rwidth+col_count,style=C$style,
                            type=C$type,
                            rule=C$rule)
                    }
                }

                col_count=rwidth+bwidth

            }
        }

        saveWorkbook(wb,fn,overwrite=TRUE)
        return(invisible(WB))
    }
)

