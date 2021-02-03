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

    list_replies=function()
    {
        op <- "replies"
        res <- private$get_paged_list(self$do_operation("replies"))
        private$init_list_objects(res, "chatMessage")
    },

    get_reply=function(message_id)
    {
        op <- file.path("replies", message_id)
        chat_message$new(self$token, self$tenant, self$do_operation(op))
    },

    send_reply=function(body, ...)
    {
        res <- self$do_operation("replies", body=call_body, http_verb="POST")
        chat_message$new(self$token, self$tenant, res)
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
))


parse_chatmsg_context <- function(x)
{
    if(is.null(x))
        stop("Unable to initialize list item object: no OData context", call.=FALSE)
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

