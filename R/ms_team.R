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
        res <- private$get_paged_list(self$do_group_operation("drives"))
        private$init_list_objects(res, "drive")
    },

    get_drive=function(drive_id=NULL)
    {
        op <- if(is.null(drive_id))
            "drive"
        else file.path("drives", drive_id)
        ms_drive$new(self$token, self$tenant, self$do_group_operation(op))
    },

    get_sharepoint_site=function()
    {
        op <- "sites/root"
        ms_site$new(self$token, self$tenant, self$do_group_operation(op))
    },

    get_group=function()
    {
        az_group$new(self$token, self$tenant, self$do_group_operation())
    },

    do_group_operation=function(op="", ...)
    {
        op <- sub("/$", "", file.path("groups", self$properties$id, op))
        call_graph_endpoint(self$token, op, ...)
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
))
