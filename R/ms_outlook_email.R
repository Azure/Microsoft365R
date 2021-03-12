#' @format An R6 object of class `ms_outlook_email`, inheriting from `ms_outlook_object`, which in turn inherits from `ms_object`.
#' @export
ms_outlook_email <- R6::R6Class("ms_outlook_email", inherit=ms_outlook_object,

public=list(

    initialize=function(token, tenant=NULL, properties=NULL)
    {
        self$type <- "email"
        private$api_type <- "messages"
        super$initialize(token, tenant, properties)
    },

    add_subject=function(subject)
    {
        self$update(subject=subject)
    },

    add_attachment=function(object)
    {
        atts <- if(is.null(self$properties$attachments))
            list()
        else self$properties$attachments
        self$update(attachments=c(atts, list(make_email_attachment(object))))
    },

    get_attachment=function(attachment_name, attachment_id) {},

    list_attachments=function() {},

    remove_attachment=function(attachment_name, attachment_id, confirm=TRUE)
    {
        self$get_attachment(attachment_name, attachment_id)$delete(confirm=confirm)
    },

    download_attachment=function(attachment_name, attachment_id, dest, overwrite=FALSE)
    {
        self$get_attachment(attachment_name, attachment_id)$download(dest, overwrite=FALSE)
    },

    set_recipients=function(to=NULL, cc=NULL, bcc=NULL)
    {
        self$update(build_email_recipients(to, cc, bcc))
    },

    send=function()
    {
        self$do_operation("send", http_verb="POST")
        invisible(NULL)
    },

    copy=function(dest) {},

    move=function(dest) {},

    reply=function(to=NULL, cc=NULL, bcc=NULL) {},

    forward=function(to=NULL, cc=NULL, bcc=NULL) {},

    print=function(...)
    {
        cat("<Outlook email>\n", sep="")
        cat("  directory id:", self$properties$id, "\n")
        cat("---\n")
        cat(self$properties$bodyPreview)
        if(nchar(self$properties$body) > nchar(self$properties$bodyPreview))
            cat("...\n")
        else cat("\n")
        cat("---\n")
        cat(format_public_methods(self))
        invisible(self)
    }
))
