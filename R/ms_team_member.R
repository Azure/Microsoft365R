ms_team_member <- R6::R6Class("ms_team_member", inherit=ms_object,

public=list(

    initialize=function(token, tenant=NULL, properties=NULL, parent_id=NULL,
                        parent_type=c("teams", "channels", "chats"))
    {
        parent_type <- match.arg(parent_type)
        if(is.null(parent_id))
            stop("Missing team/channel/conversation ID", call.=FALSE)
        self$type <- "team member"
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
        cat("<Team member '", self$properties$displayName, "'>\n", sep="")
        cat("  directory id:", self$properties$id, "\n")
        cat("  user id:", self$properties$userId, "\n")
        invisible(self)
    }
))
