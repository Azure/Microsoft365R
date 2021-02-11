#' Microsoft Teams team
#'
#' Class representing a team in Microsoft Teams.
#'
#' @docType class
#' @section Fields:
#' - `token`: The token used to authenticate with the Graph host.
#' - `tenant`: The Azure Active Directory tenant for this team.
#' - `type`: Always "team" for a team object.
#' - `properties`: The team properties.
#' @section Methods:
#' - `new(...)`: Initialize a new team object. Do not call this directly; see 'Initialization' below.
#' - `delete(confirm=TRUE)`: Delete a team. By default, ask for confirmation first.
#' - `update(...)`: Update the team metadata in Microsoft Graph.
#' - `do_operation(...)`: Carry out an arbitrary operation on the team.
#' - `sync_fields()`: Synchronise the R object with the team metadata in Microsoft Graph.
#' - `list_channels()`: List the channels for this team.
#' - `get_channel(channel_name, channel_id)`: Retrieve a channel. If the name and ID are not specified, returns the primary channel.
#' - `list_drives()`: List the drives (shared document libraries) associated with this team.
#' - `get_drive(drive_id)`: Retrieve a shared document library for this team. If the ID is not specified, this returns the default document library.
#' - `get_sharepoint_site()`: Get the SharePoint site associated with the team.
#' - `get_group()`: Get the Azure Active Directory group associated with the team.
#'
#' @section Initialization:
#' Creating new objects of this class should be done via the `get_team` and `list_teams` methods of the [ms_graph], [az_user] or [az_group] classes. Calling the `new()` method for this class only constructs the R object; it does not call the Microsoft Graph API to retrieve or create the actual team.
#'
#' @seealso
#' [ms_graph], [az_group], [ms_channel], [ms_site], [ms_drive]
#'
#' [Microsoft Graph overview](https://docs.microsoft.com/en-us/graph/overview),
#' [Microsoft Teams API reference](https://docs.microsoft.com/en-us/graph/api/resources/teams-api-overview?view=graph-rest-1.0)
#'
#' @examples
#' \dontrun{
#'
#' myteam <- get_team("my team")
#' myteam$list_channels()
#' myteam$get_channel()
#' myteam$get_drive()
#'
#' }
#' @format An R6 object of class `ms_team`, inheriting from `ms_object`.
#' @export
ms_team <- R6::R6Class("ms_team", inherit=ms_object,

public=list(

    initialize=function(token, tenant=NULL, properties=NULL)
    {
        self$type <- "team"
        private$api_type <- "teams"
        super$initialize(token, tenant, properties)
    },

    list_channels=function()
    {
        res <- private$get_paged_list(self$do_operation("channels"))
        private$init_list_objects(res, "channel", team_id=self$properties$id)
    },

    get_channel=function(channel_name=NULL, channel_id=NULL)
    {
        if(!is.null(channel_name) && is.null(channel_id))
        {
            channels <- self$list_channels()
            n <- which(sapply(channels, function(ch) ch$properties$displayName == channel_name))
            if(length(n) != 1)
                stop("Invalid channel name", call.=FALSE)
            return(channels[[n]])
        }
        op <- if(is.null(channel_name) && is.null(channel_id))
            "primaryChannel"
        else if(is.null(channel_name) && !is.null(channel_id))
            file.path("channels", channel_id)
        else stop("Do not supply both the channel name and ID", call.=FALSE)
        ms_channel$new(self$token, self$tenant, self$do_operation(op), team_id=self$properties$id)
    },

    list_drives=function()
    {
        res <- private$get_paged_list(private$do_group_operation("drives"))
        private$init_list_objects(res, "drive")
    },

    get_drive=function(drive_id=NULL)
    {
        op <- if(is.null(drive_id))
            "drive"
        else file.path("drives", drive_id)
        ms_drive$new(self$token, self$tenant, private$do_group_operation(op))
    },

    get_sharepoint_site=function()
    {
        op <- "sites/root"
        ms_site$new(self$token, self$tenant, private$do_group_operation(op))
    },

    get_group=function()
    {
        az_group$new(self$token, self$tenant, private$do_group_operation())
    },

    print=function(...)
    {
        cat("<Team '", self$properties$displayName, "'>\n", sep="")
        cat("  directory id:", self$properties$id, "\n")
        cat("  web link:", self$properties$webUrl, "\n")
        cat("  description:", self$properties$description, "\n")
        cat("---\n")
        cat(format_public_methods(self))
        invisible(self)
    }
),

private=list(

    do_group_operation=function(op="", ...)
    {
        op <- sub("/$", "", file.path("groups", self$properties$id, op))
        call_graph_endpoint(self$token, op, ...)
    }
))
