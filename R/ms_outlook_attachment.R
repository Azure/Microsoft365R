#' @format An R6 object of class `ms_outlook_attachment`, inheriting from `ms_outlook_object`, which in turn inherits from `ms_object`.
#' @export
ms_outlook_attachment <- R6::R6Class("ms_outlook_attachment", inherit=ms_outlook_object,

public=list(

    user_id=NULL,
    message_id=NULL,
    attachment_type=NULL,

    initialize=function(token, tenant=NULL, properties=NULL, user_id=NULL, message_id=NULL)
    {
        if(is.null(user_id) || is.null(message_id))
            stop("Must supply user and message IDs", call.=FALSE)
        self$type <- "email"
        self$user_id <- user_id
        self$message_id <- message_id
        self$attachment_type <- get_attachment_type(properties$`@odata.type`)
        private$api_type <- file.path("users", self$user_id, "messages", message_id, "attachments")
        super$initialize(token, tenant, properties)
    },

    sync_fields=function()
    {
        opts <- if(self$attachment_type == "file")
        {
            # don't download attachment contents if this is a file attachment
            fields <- c("id", "name", "contentType", "size", "isInline", "lastModifiedDateTime", "contentId")
            list(select=paste(fields, collapse=","))
        }
        else list()
        self$properties <- do_operation(options=opts)
        invisible(self)
    },

    download=function(dest=self$properties$name, overwrite=overwrite)
    {
        res <- self$do_operation("value", config=httr::write_disk(dest, overwrite=overwrite),
                                 http_status_handler="pass")
        if(httr::status_code(res) >= 300)
        {
            on.exit(file.remove(dest))
            httr::stop_for_status(res, paste0("complete operation. Message:\n",
                sub("\\.$", "", error_message(httr::content(res)))))
        }
        invisible(NULL)
    },

    print=function(...)
    {
        cat("<Outlook email attachment '", self$properties$name, "'>\n", sep="")
        cat("  directory id:", self$properties$id, "\n")
        cat("  attachment type:", self$attachment_type, "\n")
        invisible(self)
    }
))


get_type <- function(type)
{
    if(is_empty(type))
        stop("Unable to determine attachment type", call.=FALSE)
    switch(type,
        "#microsoft.graph.fileAttachment"="file",
        "#microsoft.graph.referenceAttachment"="link",
        "#microsoft.graph.itemAttachment"="Outlook item",
        "<unknown>"
    )
}
