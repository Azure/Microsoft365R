ms_list_item <- R6::R6Class("ms_list_item", inherit=ms_object,

public=list(

    initialize=function(token, tenant=NULL, properties=NULL)
    {
        self$type <- "list item"
        private$api_type <- "items"
        super$initialize(token, tenant, properties)
    },

    print=function(...)
    {
        cat("<Sharepoint list item '", self$properties$displayName, "'>\n", sep="")
        cat("  directory id:", self$properties$id, "\n")
        cat("---\n")
        cat(format_public_methods(self))
        invisible(self)
    }
))
