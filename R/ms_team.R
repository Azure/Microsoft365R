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
#' - `list_channels(filter=NULL, n=Inf)`: List the channels for this team.
#' - `get_channel(channel_name, channel_id)`: Retrieve a channel. If the name and ID are not specified, returns the primary channel.
#' - `create_channel(channel_name, description, membership)`: Create a new channel. Optionally, you can specify a short text description of the channel, and the type of membership: either standard or private (invitation-only).
#' - `delete_channel(channel_name, channel_id, confirm=TRUE)`: Delete a channel; by default, ask for confirmation first. You cannot delete the primary channel of a team. Note that Teams keeps track of all channels ever created, even if you delete them (you can see the deleted channels by going to the "Manage team" pane for a team, then the "Channels" tab, and expanding the "Deleted" entry); therefore, try not to create and delete channels unnecessarily.
#' - `list_drives(filter=NULL, n=Inf)`: List the drives (shared document libraries) associated with this team.
#' - `get_drive(drive_name, drive_id)`: Retrieve a shared document library for this team. If the name and ID are not specified, this returns the default document library.
#' - `get_sharepoint_site()`: Get the SharePoint site associated with the team.
#' - `get_group()`: Retrieve the Microsoft 365 group associated with the team.
#' - `list_members(filter=NULL, n=Inf)`: Retrieves the members of the team, as a list of [`ms_team_member`] objects.
#' - `get_member(name, email, id)`: Retrieve a specific member of the channel, as a `ms_team_member` object. Supply only one of the member name, email address or ID.
#'
#' @section Initialization:
#' Creating new objects of this class should be done via the `get_team` and `list_teams` methods of the [`ms_graph`], [`az_user`] or [`az_group`] classes. Calling the `new()` method for this class only constructs the R object; it does not call the Microsoft Graph API to retrieve or create the actual team.
#'
#' @section List methods:
#' All `list_*` methods have `filter` and `n` arguments to limit the number of results. The former should be an [OData expression](https://learn.microsoft.com/en-us/graph/query-parameters#filter-parameter) as a string to filter the result set on. The latter should be a number setting the maximum number of (filtered) results to return. The default values are `filter=NULL` and `n=Inf`. If `n=NULL`, the `ms_graph_pager` iterator object is returned instead to allow manual iteration over the results.
#'
#' Support in the underlying Graph API for OData queries is patchy. Not all endpoints that return lists of objects support filtering, and if they do, they may not allow all of the defined operators. If your filtering expression results in an error, you can carry out the operation without filtering and then filter the results on the client side.
#' @seealso
#' [`ms_graph`], [`az_group`], [`ms_channel`], [`ms_site`], [`ms_drive`]
#'
#' [Microsoft Graph overview](https://learn.microsoft.com/en-us/graph/overview),
#' [Microsoft Teams API reference](https://learn.microsoft.com/en-us/graph/api/resources/teams-api-overview?view=graph-rest-1.0)
#'
#' @examples
#' \dontrun{
#'
#' myteam <- get_team("my team")
#' myteam$list_channels()
#' myteam$get_channel()
#' myteam$get_drive()
#'
#' myteam$create_channel("Test channel", description="A channel for testing")
#' myteam$delete_channel("Test channel")
#'
#' # team members
#' myteam$list_members()
#' myteam$get_member("Jane Smith")
#' myteam$get_member(email="billg@mycompany.com")
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

    list_channels=function(filter=NULL, n=Inf)
    {
        private$make_basic_list("channels", filter, n, team_id=self$properties$id)
    },

    get_channel=function(channel_name=NULL, channel_id=NULL)
    {
        if(!is.null(channel_name) && is.null(channel_id))
        {
            filter <- sprintf("displayName eq '%s'", channel_name)
            channels <- self$list_channels(filter=filter)
            if(length(channels) != 1)
                stop("Invalid channel name", call.=FALSE)
            return(channels[[1]])
        }
        op <- if(is.null(channel_name) && is.null(channel_id))
            "primaryChannel"
        else if(is.null(channel_name) && !is.null(channel_id))
            file.path("channels", channel_id)
        else stop("Do not supply both the channel name and ID", call.=FALSE)
        ms_channel$new(self$token, self$tenant, self$do_operation(op), team_id=self$properties$id)
    },

    create_channel=function(channel_name, description="", membership=c("standard", "private"))
    {
        membership <- match.arg(membership)
        body <- list(
            displayName=channel_name,
            description=description,
            membershipType=membership
        )
        ms_channel$new(self$token, self$tenant, self$do_operation("channels", body=body, http_verb="POST"),
                       team_id=self$properties$id)
    },

    delete_channel=function(channel_name=NULL, channel_id=NULL, confirm=TRUE)
    {
        assert_one_arg(channel_name, channel_id, msg="Supply exactly one of channel name or ID")
        self$get_channel(channel_name, channel_id)$delete(confirm=confirm)
    },

    list_drives=function(filter=NULL, n=Inf)
    {
        self$get_group()$list_drives(filter, n)
    },

    get_drive=function(drive_name=NULL, drive_id=NULL)
    {
        if(!is.null(drive_name) && !is.null(drive_id))
            stop("Supply at most one of drive name or ID", call.=FALSE)
        if(!is.null(drive_name))
        {
            # filtering not yet supported for drives, do it in R
            drives <- self$list_drives()
            wch <- which(sapply(drives, function(drv) drv$properties$name == drive_name))
            if(length(wch) != 1)
                stop("Invalid drive name", call.=FALSE)
            return(drives[[wch]])
        }
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

    list_members=function(filter=NULL, n=Inf)
    {
        private$make_basic_list("members", filter, n, parent_id=self$properties$id, parent_type="team")
    },

    get_member=function(name=NULL, email=NULL, id=NULL)
    {
        assert_one_arg(name, email, id, msg="Supply exactly one of member name, email address, or ID")
        if(!is.null(id))
        {
            res <- self$do_operation(file.path("members", id))
            ms_team_member$new(self$token, self$tenant, res,
                parent_id=self$properties$id, parent_type="team")
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
