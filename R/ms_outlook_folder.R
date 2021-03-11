#' @format An R6 object of class `ms_outlook_folder`, inheriting from `ms_outlook_object`, which in turn inherits from `ms_object`.
#' @export
ms_outlook_folder <- R6::R6Class("ms_outlook_folder", inherit=ms_outlook_object,

public=list(

    initialize=function(token, tenant=NULL, properties=NULL)
    {
        self$type <- "mail folder"
        private$api_type <- "mailFolders"
        super$initialize(token, tenant, properties)
    },

    list_emails=function(n=50)
    {
        lst <- private$get_paged_list(self$do_operation("messages", options=list(`$top`=n)))
        private$init_list_objects(lst, default_generator=ms_outlook_email)
    },

    get_email=function(message_id)
    {
        op <- file.path("messages", message_id)
        ms_outlook_email$new(self$token, self$tenant, self$do_operation(op))
    },

    create_email=function(body="", content_type=c("text", "html"), subject="", to=NULL, cc=NULL, bcc=NULL,
                          attachments=NULL)
    {
        content_type <- match.arg(content_type)
        req_body <- c(
            list(body=build_email_body(body, content_type)),
            add_email_recipients(to, cc, bcc)
        )
        res <- ms_email$new(self$token, self$tenant, self$do_operation("messages", body=req_body, http_verb="POST"))

        for(a in attachments)
            res$add_attachments(a)
        res
    },

    delete_email=function(message_id, confirm=TRUE)
    {
        self$get_email(message_id)$delete(confirm=confirm)
    },

    send_email=function(message_id)
    {
        self$get_email(message_id)$send()
    },

    copy_email=function(message_id, dest)
    {
        self$get_email(message_id)$copy(dest)
    },

    move_email=function(message_id, dest)
    {
        self$get_email(message_id)$move(dest)
    },

    print=function(...)
    {
        cat("<Outlook folder '", self$properties$displayName, "'>\n", sep="")
        cat("  directory id:", self$properties$id, "\n")
        cat("---\n")
        cat(format_public_methods(self))
        invisible(self)
    }
))
