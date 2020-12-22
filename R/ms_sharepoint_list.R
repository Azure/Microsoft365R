#' @export
ms_sharepoint_list <- R6::R6Class("ms_sharepoint_list", inherit=ms_object,

public=list(

    initialize=function(token, tenant=NULL, properties=NULL)
    {
        self$type <- "list"
        private$api_type <- "lists"
        super$initialize(token, tenant, properties)
    },

    list_items=function(filter=NULL, select=NULL, include_metadata=FALSE, pagesize=5000)
    {
        select <- if(is.null(select))
            "fields"
        else paste0("fields(select=", paste0(select, collapse=","), ")")
        options <- list(expand=select, filter=filter, `$top`=pagesize)
        headers <- httr::add_headers(Prefer="HonorNonIndexedQueriesWarningMayFailRandomly")

        items <- self$do_operation("items", options=options, headers)
        df <- jsonlite::fromJSON(jsonlite::toJSON(private$get_paged_list(items), auto_unbox=TRUE, null="null"))
        if(!include_metadata)
            df$fields
        else df
    },

    get_column_info=function()
    {
        res <- self$do_operation(options=list(expand="columns"), simplify=TRUE)
        res$columns
    },

    do_operation=function(op="", ...)
    {
        op <- construct_path(
            "sites", self$properties$parentReference$siteId,
            "lists", self$properties$id,
            op
        )
        call_graph_endpoint(self$token, op, ...)
    },

    print=function(...)
    {
        cat("<Sharepoint list '", self$properties$displayName, "'>\n", sep="")
        cat("  directory id:", self$properties$id, "\n")
        cat("  web link:", self$properties$webUrl, "\n")
        cat("  description:", self$properties$description, "\n")
        cat("---\n")
        cat(format_public_methods(self))
        invisible(self)
    }
))
