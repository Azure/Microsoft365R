#' Teams/channel member
#'
#' Class representing a member of a team or channel (which will normally be a user in Azure Active Directory).
#'
#' @docType class
#' @section Fields:
#' - `token`: The token used to authenticate with the Graph host.
#' - `tenant`: The Azure Active Directory tenant for the parent object.
#' - `type`: One of "team member", "channel member" or "chat member" depending on the parent object.
#' - `properties`: The item properties (metadata).
#' @section Methods:
#' - `new(...)`: Initialize a new object. Do not call this directly; see 'Initialization' below.
#' - `delete(confirm=TRUE)`: Delete this member.
#' - `update(...)`: Update the member's properties (metadata) in Microsoft Graph.
#' - `do_operation(...)`: Carry out an arbitrary operation on the member.
#' - `sync_fields()`: Synchronise the R object with the member metadata in Microsoft Graph.
#' - `get_aaduser()`: Get the AAD information for the member; returns an object of class [`AzureGraph::az_user`].
#'
#' @section Initialization:
#' Creating new objects of this class should be done via the `get_member` and `list_members` methods of the [`ms_team`]and [`ms_channel`] classes. Calling the `new()` method for this class only constructs the R object; it does not call the Microsoft Graph API to retrieve or create the actual member.
#'
#' @seealso
#' [`ms_team`], [`ms_channel`]
#'
#' [Microsoft Graph overview](https://learn.microsoft.com/en-us/graph/overview),
#' [Microsoft Teams API reference](https://learn.microsoft.com/en-us/graph/api/resources/teams-api-overview?view=graph-rest-1.0)
#' @format An R6 object of class `ms_team_member`, inheriting from `ms_object`.
#' @export
ms_team_member <- R6::R6Class("ms_team_member", inherit=ms_object,

public=list(

    initialize=function(token, tenant=NULL, properties=NULL, parent_id=NULL,
                        parent_type=c("teams", "channels", "chats"))
    {
        parent_type <- match.arg(parent_type)
        if(is.null(parent_id))
            stop("Missing team/channel/conversation ID", call.=FALSE)
        self$type <- sub("s$", " member", parent_type)
        private$api_type <- file.path(parent_type, parent_id, "members")
        super$initialize(token, tenant, properties)
    },

    get_aaduser=function()
    {
        if(is.null(self$properties$userId))
            stop("Not an Azure Active Directory user identity", call.=FALSE)
        res <- call_graph_endpoint(self$token, file.path("users", self$properties$userId))
        az_user$new(self$token, self$tenant, res)
    },

    print=function(...)
    {
        type <- self$type
        substr(type, 1, 1) <- toupper(substr(type, 1, 1))
        cat("<", type, " '", self$properties$displayName, "'>\n", sep="")
        cat("  directory id:", self$properties$id, "\n")
        cat("  user id:", self$properties$userId, "\n")
        invisible(self)
    }
))
