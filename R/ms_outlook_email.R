#' @format An R6 object of class `ms_outlook_email`, inheriting from `ms_outlook_object`, which in turn inherits from `ms_object`.
#' @export
ms_outlook_email <- R6::R6Class("ms_outlook_email", inherit=ms_outlook_object,

public=list(

    initialize=function(token, tenant=NULL, properties=NULL)
    {
        self$type <- "mail"
        private$api_type <- "messages"
        super$initialize(token, tenant, properties)
    },

    add_subject=function(subject) {},

    add_attachment=function(attachment) {},

    remove_attachment=function(attachment) {},

    download_attachment=function(attachment) {},

    add_recipients=function(to=NULL, cc=NULL, bcc=NULL) {},

    clear_recipients=function() {},

    send=function() {},

    copy=function(dest) {},

    move=function(dest) {},

    reply=function(to=NULL, cc=NULL, bcc=NULL) {},

    forward=function(to=NULL, cc=NULL, bcc=NULL) {},

    print=function(...)
    {
        cat("<Outlook email>\n", sep="")
        cat("  directory id:", self$properties$id, "\n")
        cat("---\n")
        cat(format_public_methods(self))
        invisible(self)
    }
))
