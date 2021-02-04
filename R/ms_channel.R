ms_channel <- R6::R6Class("ms_channel", inherit=ms_object,

public=list(

    initialize=function(token, tenant=NULL, properties=NULL)
    {
        self$type <- "channel"
        gid <- parse_channel_weburl(properties$webUrl)
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
        lst <- private$get_paged_list(self$do_operation("messages"), n=n)
        private$init_list_objects(lst, "chatMessage")
    },

    get_message=function(message_id)
    {
        op <- file.path("messages", message_id)
        ms_chat_message$new(self$token, self$tenant, self$do_operation(op))
    },

    delete_message=function(message_id, confirm=TRUE)
    {
        self$get_message(message_id)$delete(confirm=confirm)
    },

    get_group=function()
    {
        az_group$new(self$token, self$tenant, self$do_group_operation())
    },

    do_group_operation=function(op="", ...)
    {
        gid <- parse_channel_weburl(self$properties$webUrls)
        op <- sub("/$", "", file.path("groups", gid, op))
        call_graph_endpoint(self$token, op, ...)
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
),

private=list(

    get_paged_list=function(lst, next_link_name="@odata.nextLink", value_name="value", simplify=FALSE, n=Inf)
    {
        bind_fn <- if(requireNamespace("vctrs"))
            vctrs::vec_rbind
        else base::rbind
        res <- lst[[value_name]]
        if(n <= 0) n <- Inf
        while(!is_empty(lst[[next_link_name]]) && length(res) < n)
        {
            lst <- call_graph_url(self$token, lst[[next_link_name]], simplify=simplify)
            res <- if(simplify)
                bind_fn(res, lst[[value_name]])  # this assumes all objects have the exact same fields
            else c(res, lst[[value_name]])
        }
        if(n < length(res))
            res[seq_len(n)]
        else res
    }
))


parse_channel_weburl <- function(x)
{
    if(is.null(x))
        stop("Unable to initialize team channel object: no web URL", call.=FALSE)
    httr::parse_url(x)$query$groupId
}
