ms_channel <- R6::R6Class("ms_channel", inherit=ms_object,

public=list(

    initialize=function(token, tenant=NULL, properties=NULL)
    {
        self$type <- "channel"
        gid <- parse_channel_weburl(properties[["webUrl"]])
        private$api_type <- file.path("teams", gid, "channels")
        super$initialize(token, tenant, properties)
    },

    send_message=function(body, ...)
    {
        call_body <- list(body=body, ...)
        res <- self$do_operation("messages", body=call_body, http_verb="POST")
        chat_message$new(self$token, self$tenant, res)
    },

    list_messages=function(n=50)
    {
        lst <- private$get_paged_list(self$do_operation("messages"))
        private$init_list_objects(lst, "chatMessage")
    },

    get_message=function(message_id)
    {
        op <- file.path("messages", message_id)
        chat_message$new(self$token, self$tenant, self$do_operation(op))
    },

    delete_message=function(message_id, confirm=TRUE)
    {
        self$get_message(message_id)$delete(confirm=confirm)
    },

    print=function(...)
    {
        cat("<Teams channel '", self$properties$displayName, "'>\n", sep="")
        cat("  directory id:", self$properties$id, "\n")
        cat("  web link:", self$properties$webUrl, "\n")
        cat("---\n")
        cat(format_public_methods(self))
        invisible(self)
    }
))


parse_channel_weburl <- function(x)
{
    if(is.null(x))
        stop("Unable to initialize team channel object: no web URL", call.=FALSE)
    httr::parse_url(x)$query$groupId
}
