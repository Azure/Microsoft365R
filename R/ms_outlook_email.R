#' @format An R6 object of class `ms_outlook_email`, inheriting from `ms_outlook_object`, which in turn inherits from `ms_object`.
#' @export
ms_outlook_email <- R6::R6Class("ms_outlook_email", inherit=ms_outlook_object,

public=list(

    user_id=NULL,

    initialize=function(token, tenant=NULL, properties=NULL, user_id=NULL)
    {
        if(is.null(user_id))
            stop("Must supply user ID", call.=FALSE)
        self$type <- "email"
        self$user_id <- user_id
        private$api_type <- file.path("users", self$user_id, "messages")
        super$initialize(token, tenant, properties)
    },

    set_subject=function(subject)
    {
        self$update(subject=subject)
    },

    set_recipients=function(to=NULL, cc=NULL, bcc=NULL)
    {
        if(is_empty(to) && is_empty(cc) && is_empty(bcc))
            message("Clearing all recipients")
        do.call(self$update, build_email_recipients(to, cc, bcc))
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
        if(!is_empty(self$properties$from))
            cat("  from:", format_email_recipient(self$properties$from), "\n")
        else cat("  from:\n")

        if(!self$properties$isDraft)
            cat("  sent:", format_email_date(self$properties$sentDateTime), "\n")
        else cat("  sent:\n")

        to_fmt <- sapply(self$properties$toRecipients, format_email_recipient)
        cat("  to:", paste(to_fmt, collapse=", "), "\n")

        cc_fmt <- sapply(self$properties$ccRecipients, format_email_recipient)
        if(!is_empty(cc_fmt))
            cat("  cc:", paste(cc_fmt, collapse=", "), "\n")

        bcc_fmt <- sapply(self$properties$bccRecipients, format_email_recipient)
        if(!is_empty(bcc_fmt))
            cat("  bcc:", paste(bcc_fmt, collapse=", "), "\n")

        cat("  subject:", self$properties$subject, "\n")
        cat("---\n")

        cat(self$properties$bodyPreview)
        if(nchar(self$properties$bodyPreview) >= 255)
            cat(" ...\n")
        else cat("\n")
        invisible(self)
    }
))


format_email_recipient <- function(obj)
{
    name <- obj$emailAddress$name
    addr <- obj$emailAddress$address
    name_null <- is_empty(name) || nchar(name) == 0
    addr_null <- is_empty(addr) || nchar(addr) == 0

    if(name_null && addr_null)
        "<unknown>"
    else if(name_null)
        addr
    else name
}


format_email_date <- function(datestr)
{
    date <- as.POSIXct(datestr, format="%Y-%m-%dT%H:%M:%OS", tz="UTC")
    format(date, tz="", usetz=TRUE)
}
