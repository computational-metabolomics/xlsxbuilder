#' xlsx_builder
#'
#' data that goes somewhere in an excel sheet
#' @import openxlsx
#' @export xlsx_builder
xlsx_builder<-setClass(
    "xlsx_builder",
    slots=c(hidden='character')
)


#' @export
setMethod(f="$",
    signature=c("xlsx_builder"),
    definition=function(x,name)
    {
        s=slotNames(x)

        if ((name %in% s) && (!name %in% x@hidden)) {
            value=slot(x,name)
            return(value)
        } else {
            stop(paste0('"',name,'" is not a valid slot'))
        }
    }
)

#' @export
setMethod(f="$<-",
    signature(x='xlsx_builder'),
    definition=function(x,name,value) {
        s=slotNames(x)
        if ((name %in% s) && (!name %in% x@hidden)) {
            slot(x,name)=value
            return(x)
        } else {
            stop(paste0('"',name,'" is not a valid slot'))
        }
    }
)
