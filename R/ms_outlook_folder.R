#' @format An R6 object of class `ms_outlook_folder`, inheriting from `ms_outlook_object`, which in turn inherits from `ms_object`.
#' @export
ms_outlook_folder <- R6::R6Class("ms_outlook_folder", inherit=ms_outlook_object,

public=list(

    user_id=NULL,

    initialize=function(token, tenant=NULL, properties=NULL, user_id=NULL)
    {
        if(is.null(user_id))
            stop("Must supply user ID", call.=FALSE)
        self$type <- "mail folder"
        self$user_id <- user_id
        private$api_type <- file.path("users", self$user_id, "mailFolders")
        super$initialize(token, tenant, properties)
    },

    list_emails=function(n=50)
    {
        lst <- private$get_paged_list(self$do_operation("messages", options=list(`$top`=n)))
        private$init_list_objects(lst, default_generator=ms_outlook_email, user_id=self$user_id)
    },

    get_email=function(message_id)
    {
        op <- file.path("messages", message_id)
        ms_outlook_email$new(self$token, self$tenant, self$do_operation(op), user_id=self$user_id)
    },

    create_email=function(body="", content_type=c("text", "html"), subject="", to=NULL, cc=NULL, bcc=NULL,
                          reply_to=NULL, attachments=NULL, send_now=FALSE)
    {
        content_type <- match.arg(content_type)
        req <- build_email_request(body, content_type, attachments, subject, to, cc, bcc)
        res <- ms_outlook_email$new(self$token, self$tenant,
            self$do_operation("messages", body=req, http_verb="POST"), user_id=self$user_id)

        if(send_now)
        {
            res$send()
            return(invisible(NULL))  # sent email no longer has any link to draft, despite immutable IDs (!)
        }

        res
    },

    delete_email=function(message_id, confirm=TRUE)
    {
        self$get_email(message_id)$delete(confirm=confirm)
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
