ms_chat_message <- R6::R6Class("ms_chat_message", inherit=ms_object,

public=list(

    initialize=function(token, tenant=NULL, properties=NULL)
    {
        self$type <- "Teams message"
        context <- parse_chatmsg_context(properties[["fields@odata.context"]])
        private$api_type <- if(!is.null(context$team_id))
            file.path("teams", context$team_id, "channels", context$channel_id, "messages")
        else file.path("users", context$user_id, "chats", context$chat_id, "messages")
        super$initialize(token, tenant, properties)
    },

    list_replies=function(n=50)
    {
        op <- "replies"
        res <- private$get_paged_list(self$do_operation("replies"), n=n)
        private$init_list_objects(res, "chatMessage")
    },

    get_reply=function(message_id)
    {
        op <- file.path("replies", message_id)
        ms_chat_message$new(self$token, self$tenant, self$do_operation(op))
    },

    send_reply=function(body, ...)
    {
        res <- self$do_operation("replies", body=call_body, http_verb="POST")
        ms_chat_message$new(self$token, self$tenant, res)
    },

    delete_reply=function(message_id, confirm=TRUE)
    {
        self$get_reply(message_id)$delete(confirm=confirm)
    },

    print=function(...)
    {
        cat("<Teams message>\n", sep="")
        cat("  directory id:", self$properties$id, "\n")
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
        if(n < Inf)
            res[seq_len(n)]
        else res
    }
))


parse_chatmsg_context <- function(x)
{
    if(is.null(x))
        stop("Unable to initialize Teams message object: no OData context", call.=FALSE)
    # is this a channel or chat msg?
    if(grepl("^.+#teams\\('", x))
    {
        x <- sub("^.+#teams\\('", "", x)
        tid <- utils::URLdecode(sub("'\\).+$", "", x))
        x <- sub("^.+channels\\('", "", x)
        cid <- sub("'\\).+", "", x)
        list(team_id=tid, channel_id=cid)
    }
    else
    {
        x <- sub("^.+#users\\('", "", x)
        uid <- utils::URLdecode(sub("'\\).+$", "", x))
        x <- sub("^.+chats\\('", "", x)
        cid <- sub("'\\).+", "", x)
        list(user_id=uid, chat_id=cid)
    }
}

