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
            for (m in 1:nt) {
                # for each table

                hheight=nrow(WB$sheets[[k]]$tables[[m]]$col_header$data)
                rwidth=ncol(WB$sheets[[k]]$tables[[m]]$row_header$data)
                rheight=nrow(WB$sheets[[k]]$tables[[m]]$row_header$data)
                hwidth=ncol(WB$sheets[[k]]$tables[[m]]$col_header$data)
                bheight=nrow(WB$sheets[[k]]$tables[[m]]$body$data)
                bwidth=ncol(WB$sheets[[k]]$tables[[m]]$body$data)

                # put the row header
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

                # put the column header
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


                # put the data
                writeData(wb=wb,sheet=WB$sheets[[k]]$name,
                    startRow = row_count+hheight+1,
                    startCol= col_count+rwidth+1,
                    colNames=FALSE,borders = WB$sheets[[k]]$tables[[m]]$body$border,
                    borderStyle = WB$sheets[[k]]$tables[[m]]$body$border_style,
                    x=WB$sheets[[k]]$tables[[m]]$body$data)
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

                col_count=rwidth+bwidth

            }
        }

        saveWorkbook(wb,fn,overwrite=TRUE)
        return(invisible(WB))
    }
)

