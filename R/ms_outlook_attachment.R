#' Outlook mail attachment
#'
#' Class representing an attachment in Outlook.
#'
#' @docType class
#' @section Fields:
#' - `token`: The token used to authenticate with the Graph host.
#' - `tenant`: The Azure Active Directory tenant for the email account.
#' - `type`: always "attachment" for an attachment.
#' - `properties`: The attachment properties (metadata).
#' @section Methods:
#' - `new(...)`: Initialize a new object. Do not call this directly; see 'Initialization' below.
#' - `delete(confirm=TRUE)`: Delete this attachment. By default, ask for confirmation first.
#' - `update(...)`: Update the attachment's properties (metadata) in Microsoft Graph.
#' - `do_operation(...)`: Carry out an arbitrary operation on the attachment.
#' - `sync_fields()`: Synchronise the R object with the attachment metadata in Microsoft Graph. This method does _not_ transfer the attachment content for a file attachment.
#' - `download(dest, overwrite)`: For a file attachment, downloads the content to a file. The default destination filename is the name of the attachment.
#'
#' @section Initialization:
#' Creating new objects of this class should be done via the the `get_attachment()`, `list_attachments()` or `create_attachment()` methods [`ms_outlook_email`] class. Calling the `new()` method for this class only constructs the R object; it does not call the Microsoft Graph API to retrieve or create the actual attachment.
#'
#' In general, you should not need to interact directly with this class, as the `ms_outlook_email` class exposes convenience methods for working with attachments. The only exception is to download an attachment in a reliable way (not involving the attachment name); see the example below.
#'
#' @seealso
#' [`ms_outlook`], [`ms_outlook_email`]
#'
#' [Microsoft Graph overview](https://learn.microsoft.com/en-us/graph/overview),
#' [Outlook API reference](https://learn.microsoft.com/en-us/graph/api/resources/mail-api-overview?view=graph-rest-1.0)
#'
#' @examples
#' \dontrun{
#'
#' outl <- get_personal_outlook()
#'
#' em <- outl$get_inbox$get_email("email_id")
#'
#' # download the first attachment in an email
#' atts <- em$list_attachments()
#' atts[[1]]$download()
#'
#' }
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

    download=function(dest=self$properties$name, overwrite=FALSE)
    {
        res <- self$do_operation("$value", config=httr::write_disk(dest, overwrite=overwrite),
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
        if(self$attachment_type == "file")
            cat("  attachment size:", self$properties$size, "\n")
        invisible(self)
    }
))


get_attachment_type <- function(type)
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
