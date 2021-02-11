#' Teams channel
#'
#' Class representing a Microsoft Teams channel.
#'
#' @docType class
#' @section Fields:
#' - `token`: The token used to authenticate with the Graph host.
#' - `tenant`: The Azure Active Directory tenant for the parent drive.
#' - `type`: Always "channel" for a channel object
#' - `team_id`: The ID of the parent team.
#' - `properties`: The item properties (metadata).
#' @section Methods:
#' - `new(...)`: Initialize a new object. Do not call this directly; see 'Initialization' below.
#' - `delete(confirm=TRUE)`: Delete this channel. By default, ask for confirmation first.
#' - `update(...)`: Update the channel's properties (metadata) in Microsoft Graph.
#' - `do_operation(...)`: Carry out an arbitrary operation on the channel.
#' - `sync_fields()`: Synchronise the R object with the channel metadata in Microsoft Graph.
#' - `send_message(body, content_type, attachments)`: Sends a new message to the channel. See below.
#' - `list_messages(n=50)`: Retrieves the messages in the channel. By default, this is limited to the 50 most recent messages; set the `n` argument to change this.
#' - `get_message(message_id)`: Retrieves a specific message in the channel.
#' - `delete_message(message_id, confirm=TRUE)`: Deletes a message. By default, ask for confirmation first.
#' - `list_files()`: List the files for the channel.
#' - `upload_file()`: Uploads a file to the channel.
#' - `download_file()`: Downloads a file from the channel.
#'
#' @section Initialization:
#' Creating new objects of this class should be done via the `get_channel` and `list_channels` methods of the [ms_team] class. Calling the `new()` method for this class only constructs the R object; it does not call the Microsoft Graph API to retrieve or create the actual channel.
#'
#' @section Messaging:
#' To send a message to a channel, use the `send_message()` method. This has arguments:
#' - `body`: The body of the message. This should be a character vector, which will be concatenated into a single string with newline separators. The body can be either plain text or HTML formatted.
#' - `content_type`: Either "text" (the default) or "html".
#' - `attachments`: Optional vector of filenames.
#'
#' Note that message attachments are actually uploaded to the channel's file listing (a directory in the team's primary shared document folder). Support for attachments is somewhat experimental, so if you want to be sure that it works, upload the file separately using the `upload_file()` method.
#'
#' @seealso
#' [ms_team], [ms_chat_message]
#'
#' [Microsoft Graph overview](https://docs.microsoft.com/en-us/graph/overview),
#' [Microsoft Teams API reference](https://docs.microsoft.com/en-us/graph/api/resources/teams-api-overview?view=graph-rest-1.0)
#'
#' @examples
#' \dontrun{
#'
#' myteam <- team("my team")
#' myteam$list_channels()
#'
#' chan <- myteam$get_channel()
#' chan$list_messages()
#' chan$send_message("hello from R")
#'
#' # a multi-line message with an attachment
#' msg_text <- c(
#'     "message line 1",
#'     "message line 2",
#'     "message line 3"
#' )
#' chan$send_message(msg_text, attachments="myfile.csv")
#'
#' chan$upload_file("mydocument.docx")
#'
#' }
#' @format An R6 object of class `ms_channel`, inheriting from `ms_object`.
#' @export
ms_channel <- R6::R6Class("ms_channel", inherit=ms_object,

public=list(

    team_id=NULL,

    initialize=function(token, tenant=NULL, properties=NULL, team_id=NULL)
    {
        self$type <- "channel"
        self$team_id <- team_id
        private$api_type <- file.path("teams", self$team_id, "channels")
        super$initialize(token, tenant, properties)
    },

    send_message=function(body, content_type=c("text", "html"), attachments=NULL)
    {
        content_type <- match.arg(content_type)
        call_body <- build_chatmessage_body(self, body, content_type, attachments)
        res <- self$do_operation("messages", body=call_body, http_verb="POST")
        ms_chat_message$new(self$token, self$tenant, res)
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

    list_files=function(path="", ...)
    {
        path <- sub("/$", "", file.path(self$properties$displayName, path))
        private$get_drive()$list_files(path, ...)
    },

    download_file=function(src, dest=basename(src), ...)
    {
        src <- file.path(self$properties$displayName, src)
        private$get_drive()$download_file(src, dest, ...)
    },

    upload_file=function(src, dest, ...)
    {
        dest <- file.path(self$properties$displayName, dest)
        private$get_drive()$upload_file(src, dest, ...)
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

    get_drive=function(drive_id=NULL)
    {
        op <- if(is.null(drive_id))
            "drive"
        else file.path("drives", drive_id)
        ms_drive$new(self$token, self$tenant, private$do_group_operation(op))
    },

    get_group=function()
    {
        az_group$new(self$token, self$tenant, private$do_group_operation())
    },

    do_group_operation=function(op="", ...)
    {
        op <- sub("/$", "", file.path("groups", self$team_id, op))
        call_graph_endpoint(self$token, op, ...)
    }
))
