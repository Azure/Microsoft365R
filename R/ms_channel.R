ms_channel <- R6::R6Class("ms_channel", inherit=ms_object,

public=list(

    initialize=function(token, tenant=NULL, properties=NULL)
    {
        self$type <- "channel"
        gid <- parse_channel_weburl(properties[["webUrl"]])
        private$api_type <- file.path("teams", gid, "channels")
        super$initialize(token, tenant, properties)
    },

    post_message=function() {},

    print=function(...)
    {
        cat("<Teams channel '", self$properties$displayName, "'>\n", sep="")
        cat("  directory id:", self$properties$id, "\n")
        cat("  web link:", self$properties$webUrl, "\n")
        cat("---\n")
        cat(format_public_methods(self))
        invisible(self)
    }
))


parse_channel_weburl <- function(x)
{
    if(is.null(x))
        stop("Unable to initialize team channel object: no web URL", call.=FALSE)
    httr::parse_url(x)$query$groupId
}
