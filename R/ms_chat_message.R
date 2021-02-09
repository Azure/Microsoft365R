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

    send_reply=function(body, content_type=c("text", "html"), attachments=NULL)
    {
        private$assert_not_nested_reply()
        content_type <- match.arg(content_type)
        call_body <- build_chatmessage_body(private$get_channel(), body, content_type, attachments)
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
        if(!is_empty(self$properties$replyToId))
            cat("  in-reply-to:", self$properties$replyToId, "\n")
        cat("---\n")
        cat(format_public_methods(self))
        invisible(self)
    }
),

private=list(

    get_channel=function()
    {
        channel <- self$properties$channelIdentity
        ms_channel$new(self$token, self$tenant, list(id=channel$channelId), team_id=channel$teamId)$sync_fields()
    },

    assert_not_nested_reply=function()
    {
        stopifnot("Nested replies not allowed in Teams channels"=is.null(self$properties$replyToId))
    }
))


build_chatmessage_body <- function(channel, body, content_type, attachments)
{
    call_body <- list(body=list(content=paste(body, collapse="\n"), contentType=content_type))
    if(!is_empty(attachments))
    {
        call_body$attachments <- lapply(attachments, function(f)
        {
            att <- channel$upload_file(f, dest=basename(f))
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
    call_body
}
