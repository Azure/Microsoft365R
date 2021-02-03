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
        private$init_list_objects(res, "channel")
    },

    get_channel=function(channel_id=NULL)
    {
        op <- if(is.null(channel_id))
            "primaryChannel"
        else file.path("channels", drive_id)
        ms_channel$new(self$token, self$tenant, self$do_operation(op))
    },

    list_drives=function()
    {
        op <- private$group_path("drives")
        lst <- private$get_paged_list(call_graph_endpoint(self$token, op))
        private$init_list_objects(lst, "drive")
    },

    get_drive=function(drive_id=NULL)
    {
        op <- if(is.null(drive_id))
            private$group_path("drive")
        else private$group_path("drives", drive_id)
        ms_drive$new(self$token, self$tenant, call_graph_endpoint(self$token, op))
    },

    get_sharepoint_site=function()
    {
        op <- private$group_path("sites/root")
        ms_site$new(self$token, self$tenant, call_graph_endpoint(self$token, op))
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

    group_path=function(...)
    {
        file.path("groups", self$properties$id, ...)
    }
))
