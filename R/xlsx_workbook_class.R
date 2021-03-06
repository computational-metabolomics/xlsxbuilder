#' xlsx_workbook
#'
#' a workbook combines multiple sheets.
#' @import openxlsx
#' @include xlsx_sheet_class.R
#' @export xlsx_workbook
xlsx_workbook<-setClass(
    "xlsx_workbook",
    contains='xlsx_builder',
    slots=c(sheets='list',overwrite='logical',padding='numeric'),
    prototype=list(overwrite=FALSE,padding=2)
)

#' @export
setMethod(f="xlsx_write",
    signature=c("xlsx_workbook",'character'),
    definition=function(WB,fn) {

        # prepare workbook
        wb=createWorkbook()

        # add sheets
        ns=length(WB$sheets)

        # keep row_counts of sheets by name
        rn=list()

        for (k in 1:ns) {

            # add the sheet if it doesnt exist
            if (!(WB$sheets[[k]]$name %in% wb$sheet_names)) {
                addWorksheet(wb,WB$sheets[[k]]$name)
                row_count = 0 # current row number
            } else {
                # add the tables to the bottom of the named sheet
                row_count=rn[[WB$sheets[[k]]$name]]
            }

            col_count = 0 # current column number

            # for each sheet
            nt=length(WB$sheets[[k]]$tables)

            # count the number of header rows
            head_count=numeric(nt)
            data_count=numeric(nt)
            for(m in 1:nt) {
                head_count[m]=nrow(WB$sheets[[k]]$tables[[m]]$col_header$data)
            }

            # pad the headers of smaller tables with spaces to make them the same size as the largest one
            for(m in 1:nt) {
                if (head_count[m]<max(head_count)) {
                    temp=matrix(NA,nrow=max(head_count)-head_count[m],ncol=ncol(WB$sheets[[k]]$tables[[m]]$col_header$data))
                    temp=as.data.frame(temp)
                    colnames(temp)=colnames(WB$sheets[[k]]$tables[[m]]$col_header$data)
                    WB$sheets[[k]]$tables[[m]]$col_header$data=rbind(temp,WB$sheets[[k]]$tables[[m]]$col_header$data)
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
                        startCol= col_count + 1 ,
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

                col_count=col_count+rwidth+bwidth+ WB$sheets[[k]]$padding

            }

            row_count=hheight+bheight+WB$padding

            rn[[WB$sheets[[k]]$name]]=row_count
        }

        print('Writing file...')
        saveWorkbook(wb,fn,overwrite=TRUE)
        print('Success!')
        return(invisible(WB))
    }
)

#' @export
setMethod("+",
    signature(e1='xlsx_sheet',e2='xlsx_sheet'),
    definition=function(e1,e2) {
        OUT=xlsx_workbook(sheets=c(e1,e2))
        return(OUT)
    }
)

#' @export
setMethod("+",
    signature(e1='xlsx_workbook',e2='xlsx_sheet'),
    definition=function(e1,e2) {
        e1$sheets=c(e1$sheets,e2)
        return(e1)
    }
)

#' @export
setMethod("+",
    signature(e1='xlsx_sheet',e2='xlsx_workbook'),
    definition=function(e1,e2) {
        e2$tables=c(e1,e2$tables)
        return(e2)
    }
)

