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
#' - `list_messages(filter=NULL, n=50)`: Retrieves the messages in the channel. By default, this is limited to the 50 most recent messages; set the `n` argument to change this.
#' - `get_message(message_id)`: Retrieves a specific message in the channel.
#' - `delete_message(message_id, confirm=TRUE)`: Deletes a message. Currently the Graph API does not support deleting Teams messages, so this method is disabled.
#' - `list_files()`: List the files for the channel. See [`ms_drive`] for the arguments available for this and the file upload/download methods.
#' - `upload_file()`: Uploads a file to the channel.
#' - `download_file()`: Downloads a file from the channel.
#' - `get_folder()`: Retrieves the files folder for the channel, as a [`ms_drive_item`] object.
#' - `list_members(filter=NULL, n=Inf)`: Retrieves the members of the channel, as a list of [`ms_team_member`] objects.
#' - `get_member(name, email, id)`: Retrieve a specific member of the channel, as a `ms_team_member` object. Supply only one of the member name, email address or ID.
#'
#' @section Initialization:
#' Creating new objects of this class should be done via the `get_channel` and `list_channels` methods of the [`ms_team`] class. Calling the `new()` method for this class only constructs the R object; it does not call the Microsoft Graph API to retrieve or create the actual channel.
#'
#' @section Messaging:
#' To send a message to a channel, use the `send_message()` method. This has arguments:
#' - `body`: The body of the message. This should be a character vector, which will be concatenated into a single string with newline separators. The body can be either plain text or HTML formatted.
#' - `content_type`: Either "text" (the default) or "html".
#' - `attachments`: Optional vector of filenames.
#' - `inline`: Optional vector of image filenames that will be inserted into the body of the message. The images must be PNG or JPEG, and the `content_type` argument must be "html" to include inline content.
#' - `mentions`: Optional vector of @mentions that will be inserted into the body of the message. This should be either an object of one of the following classes, or a list of the same: [`az_user`], [`ms_team`], [`ms_channel`], [`ms_team_member`]. The `content_type` argument must be "html" to include mentions.
#'
#' Note that message attachments are actually uploaded to the channel's file listing (a directory in the team's primary shared document folder). Support for attachments is somewhat experimental, so if you want to be sure that it works, upload the file separately using the `upload_file()` method.
#'
#' @section List methods:
#' All `list_*` methods have `filter` and `n` arguments to limit the number of results. The former should be an [OData expression](https://learn.microsoft.com/en-us/graph/query-parameters#filter-parameter) as a string to filter the result set on. The latter should be a number setting the maximum number of (filtered) results to return. The default values are `filter=NULL` and `n=Inf`. If `n=NULL`, the `ms_graph_pager` iterator object is returned instead to allow manual iteration over the results.
#'
#' Support in the underlying Graph API for OData queries is patchy. Not all endpoints that return lists of objects support filtering, and if they do, they may not allow all of the defined operators. If your filtering expression results in an error, you can carry out the operation without filtering and then filter the results on the client side.
#' @seealso
#' [`ms_team`], [`ms_drive`], [`ms_chat_message`]
#'
#' [Microsoft Graph overview](https://learn.microsoft.com/en-us/graph/overview),
#' [Microsoft Teams API reference](https://learn.microsoft.com/en-us/graph/api/resources/teams-api-overview?view=graph-rest-1.0)
#'
#' @examples
#' \dontrun{
#'
#' myteam <- get_team("my team")
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
#' # sending an inline image
#' chan$send_message("", content_type="html", inline="graph.png")
#'
#' # channel members
#' chan$list_members()
#' jane <- chan$get_member("Jane Smith")
#' bill <- chan$get_member(email="billg@mycompany.com")
#'
#' # mentioning a team member
#' chan$send_message("Here is a message", content_type="html", mentions=jane)
#'
#' # mentioning 2 or more members: use a list
#' chan$send_message("Here is another message", content_type="html",
#'     mentions=list(jane, bill))
#'
#' # mentioning an entire channel or team
#' chan$send_message("FYI to channel", content_type="html", mentions=chan)
#' chan$send_message("FYI to everyone", content_type="html", mentions=myteam)
#'
#' chan$list_files()
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
        if(is.null(team_id))
            stop("Missing team ID", call.=FALSE)
        self$type <- "channel"
        self$team_id <- team_id
        private$api_type <- file.path("teams", self$team_id, "channels")
        super$initialize(token, tenant, properties)
    },

    send_message=function(body, content_type=c("text", "html"), attachments=NULL, inline=NULL, mentions=NULL)
    {
        content_type <- match.arg(content_type)
        call_body <- build_chatmessage_body(self, body, content_type, attachments, inline, mentions)
        res <- self$do_operation("messages", body=call_body, http_verb="POST")
        ms_chat_message$new(self$token, self$tenant, res)
    },

    list_messages=function(filter=NULL, n=50)
    {
        private$make_basic_list("messages", filter, n)
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
        self$get_folder()$list_files(path, ...)
    },

    download_file=function(src, dest=basename(src), ...)
    {
        self$get_folder()$get_item(src)$download(dest, ...)
    },

    upload_file=function(src, dest=basename(src), ...)
    {
        self$get_folder()$upload(src, dest, ...)
    },

    get_folder=function()
    {
        if(is.null(private$folder))
            private$folder <- ms_drive_item$new(self$token, self$tenant, self$do_operation("filesFolder"))
        private$folder
    },

    list_members=function(filter=NULL, n=Inf)
    {
        private$make_basic_list("members", filter, n, parent_id=self$properties$id, parent_type="channel")
    },

    get_member=function(name=NULL, email=NULL, id=NULL)
    {
        assert_one_arg(name, email, id, msg="Supply exactly one of member name, email address, or ID")
        if(!is.null(id))
        {
            res <- self$do_operation(file.path("members", id))
            ms_team_member$new(self$token, self$tenant, res,
                parent_id=self$properties$id, parent_type="channel")
        }
        else
        {
            filter <- if(!is.null(name))
                sprintf("displayName eq '%s'", name)
            else sprintf("microsoft.graph.aadUserConversationMember/email eq '%s'", email)
            res <- self$list_members(filter=filter)
            if(length(res) != 1)
                stop("Invalid name or email address", call.=FALSE)
            res[[1]]
        }
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

    folder=NULL
))
