#' Teams chat message
#'
#' Class representing a message in a Teams channel or chat.
#'
#' @docType class
#' @section Fields:
#' - `token`: The token used to authenticate with the Graph host.
#' - `tenant`: The Azure Active Directory tenant for the parent drive.
#' - `type`: Always "Teams message" for a chat message object.
#' - `properties`: The item properties (metadata).
#' @section Methods:
#' - `new(...)`: Initialize a new object. Do not call this directly; see 'Initialization' below.
#' - `delete(confirm=TRUE)`: Delete this item. By default, ask for confirmation first.
#' - `update(...)`: Update the item's properties (metadata) in Microsoft Graph. To update the list _data_, update the `fields` property. See the examples below.
#' - `do_operation(...)`: Carry out an arbitrary operation on the item.
#' - `sync_fields()`: Synchronise the R object with the item metadata in Microsoft Graph.
#' - `send_reply(body, content_type, attachments)`: Sends a reply to the message. See below.
#' - `list_replies(n=50)`: List the replies to this message. By default, this is limited to the 50 most recent replies; set the `n` argument to change this.
#' - `get_reply(message_id)`: Retrieves a specific reply to the message.
#' - `delete_reply(message_id, confirm=TRUE)`: Deletes a reply to the message. By default, ask for confirmation first.
#'
#' @section Initialization:
#' Creating new objects of this class should be done via the `get_message` and `list_messages` method of the [ms_team] class. Calling the `new()` method for this class only constructs the R object; it does not call the Microsoft Graph API to retrieve or create the actual message.
#'
#' @section Replying to a message:
#' To reply to a message, use the `send_reply()` method. This has arguments:
#' - `body`: The body of the message. This should be a character vector, which will be concatenated into a single string with newline separators. The body can be either plain text or HTML formatted.
#' - `content_type`: Either "text" (the default) or "html".
#' - `attachments`: Optional vector of filenames.
#'
#' Teams channels don't support nested replies, so replying to a reply will fail.
#'
#' Note that message attachments are actually uploaded to the channel's file listing (a directory in the team's primary shared document folder). Support for attachments is somewhat experimental, so if you want to be sure that it works, upload the file separately using the `upload_file()` method.
#'
#' @seealso
#' [ms_team], [ms_channel]
#'
#' [Microsoft Graph overview](https://docs.microsoft.com/en-us/graph/overview),
#' [Microsoft Teams API reference](https://docs.microsoft.com/en-us/graph/api/resources/teams-api-overview?view=graph-rest-1.0)
#'
#' @examples
#' \dontrun{
#'
#' myteam <- get_team("my team")
#'
#' chan <- myteam$get_channel()
#' msg <- chan$list_messages()[[1]]
#' msg$list_replies()
#' msg$send_reply("Reply from R")
#'
#' }
#' @format An R6 object of class `ms_chat_message`, inheriting from `ms_object`.
#' @export
ms_chat_message <- R6::R6Class("ms_chat_message", inherit=ms_object,

public=list(

    initialize=function(token, tenant=NULL, properties=NULL)
    {
        self$type <- "Teams message"
        parent <- properties$channelIdentity
        private$api_type <- file.path("teams", parent[[1]], "channels", parent[[2]], "messages")
        super$initialize(token, tenant, properties)
    },

    send_reply=function(body, content_type=c("text", "html"), attachments=NULL)
    {
        private$assert_not_nested_reply()
        content_type <- match.arg(content_type)
        call_body <- build_chatmessage_body(private$get_channel(), body, content_type, attachments)
        res <- self$do_operation("replies", body=call_body, http_verb="POST")
        ms_chat_message$new(self$token, self$tenant, res)
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
