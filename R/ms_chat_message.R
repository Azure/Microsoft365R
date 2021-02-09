ms_chat_message <- R6::R6Class("ms_chat_message", inherit=ms_object,

public=list(

    initialize=function(token, tenant=NULL, properties=NULL)
    {
        self$type <- "Teams message"
        parent <- properties$channelIdentity
        private$api_type <- file.path("teams", parent[[1]], "channels", parent[[2]], "messages")
        super$initialize(token, tenant, properties)
    },

    list_replies=function(n=50)
    {
        private$assert_not_nested_reply()
        res <- private$get_paged_list(self$do_operation("replies"), n=n)
        private$init_list_objects(res, "chatMessage")
    },

    get_reply=function(message_id)
    {
        private$assert_not_nested_reply()
        op <- file.path("replies", message_id)
        ms_chat_message$new(self$token, self$tenant, self$do_operation(op))
    },

    send_reply=function(body, content_type=c("text", "html"), attachments=NULL, ...)
    {
        private$assert_not_nested_reply()
        content_type <- match.arg(content_type)
        call_body <- list(body=list(content=paste(body, collapse="\n"), contentType=content_type), ...)
        if(!is_empty(attachments))
        {
            chan <- private$get_channel()
            call_body$attachments <- lapply(attachments, function(f)
            {
                att <- chan$upload_file(f, dest=basename(f))
                list(
                    id=uuid::UUIDgenerate(),
                    name=att$properties$name,
                    contentUrl=att$properties$webUrl,
                    contentType="reference"
                )
            })
            att_tags <- lapply(call_body$attachments,
                function(att) paste0('<attachment id="', att$id, '"></attachment>'))
            call_body$body$content <- paste(call_body$body$content, paste(att_tags, collapse=""))
        }

        res <- self$do_operation("replies", body=call_body, http_verb="POST")
        ms_chat_message$new(self$token, self$tenant, res)
    },

    delete_reply=function(message_id, confirm=TRUE)
    {
        self$get_reply(message_id)$delete(confirm=confirm)
    },

    print=function(...)
    {
        parent <- self$properties$channelIdentity
        cat("<Teams message>\n", sep="")
        cat("  directory id:", self$properties$id, "\n")
        cat("  team:", parent[[1]], "\n")
        cat("  channel:", parent[[2]], "\n")
        cat("---\n")
        cat(format_public_methods(self))
        invisible(self)
    }
),

private=list(

    get_channel=function()
    {
        channel <- self$properties$channelIdentity
        ms_channel$new(self$token, self$tenant, list(id=channel$channelId), team_id=channel$teamId)
    },

    assert_not_nested_reply=function()
    {
        stopifnot("Nested replies not allowed in Teams channels"=is.null(self$properties$replyToId))
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
