ms_chat_message <- R6::R6Class("ms_chat_message", inherit=ms_object,

public=list(

    parent=NULL,

    initialize=function(token, tenant=NULL, properties=NULL, parent=NULL)
    {
        self$type <- "Teams message"
        self$parent <- parent
        private$api_type <- file.path("teams", parent$team_id, "channels", parent$channel_id, "messages")
        super$initialize(token, tenant, properties)
    },

    list_replies=function(n=50)
    {
        private$assert_not_nested_reply()
        op <- "replies"
        parent <- c(self$parent, message_id=self$properties$id)
        res <- private$get_paged_list(self$do_operation("replies"), n=n)
        private$init_list_objects(res, "chatMessage", parent=parent)
    },

    get_reply=function(message_id)
    {
        private$assert_not_nested_reply()
        op <- file.path("replies", message_id)
        parent <- c(self$parent, message_id=self$properties$id)
        ms_chat_message$new(self$token, self$tenant, self$do_operation(op), parent=parent)
    },

    send_reply=function(body, content_type=c("text", "html"), ...)
    {
        private$assert_not_nested_reply()
        content_type <- match.arg(content_type)
        call_body <- list(body=list(content=body, contentType=content_type), ...)
        parent <- c(self$parent, message_id=self$properties$id)
        res <- self$do_operation("replies", body=call_body, http_verb="POST")
        ms_chat_message$new(self$token, self$tenant, res, parent=parent)
    },

    delete_reply=function(message_id, confirm=TRUE)
    {
        self$get_reply(message_id)$delete(confirm=confirm)
    },

    print=function(...)
    {
        cat("<Teams message>\n", sep="")
        cat("  directory id:", self$properties$id, "\n")
        cat("  team:", self$parent$team_id, "\n")
        cat("  channel:", self$parent$channel_id, "\n")
        cat("---\n")
        cat(format_public_methods(self))
        invisible(self)
    }
),

private=list(

    assert_not_nested_reply=function()
    {
        stopifnot("Nested replies not allowed in Teams channels"=is.null(self$parent$message_id))
    }

#     get_paged_list=function(lst, next_link_name="@odata.nextLink", value_name="value", simplify=FALSE, n=Inf)
#     {
#         res <- lst[[value_name]]
#         if(n <= 0) n <- Inf
#         while(!is_empty(lst[[next_link_name]]) && length(res) < n)
#         {
#             lst <- call_graph_url(self$token, lst[[next_link_name]], simplify=simplify)
#             res <- if(simplify)
#                 vctrs::vec_rbind(res, lst[[value_name]])
#             else c(res, lst[[value_name]])
#         }
#         if(n < length(res))
#             res[seq_len(n)]
#         else res
#     }
))
