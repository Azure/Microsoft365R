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
