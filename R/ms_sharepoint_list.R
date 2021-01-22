#' Sharepoint list
#'
#' Class representing a list in a SharePoint site.
#'
#' @docType class
#' @section Fields:
#' - `token`: The token used to authenticate with the Graph host.
#' - `tenant`: The Azure Active Directory tenant for the parent drive.
#' - `type`: always "list" for a SharePoint list object.
#' - `properties`: The item properties (metadata).
#' @section Methods:
#' - `new(...)`: Initialize a new object. Do not call this directly; see 'Initialization' below.
#' - `delete(confirm=TRUE)`: Delete this item. By default, ask for confirmation first.
#' - `update(...)`: Update the item's properties in Microsoft Graph.
#' - `do_operation(...)`: Carry out an arbitrary operation on the item.
#' - `sync_fields()`: Synchronise the R object with the item metadata in Microsoft Graph.
#' - `list_items(filter, select, all_metadata, pagesize)`: Queries the list and returns items as a data frame. See 'List querying below'.
#' - `get_column_info()`: Return a data frame containing metadata on the columns (fields) in the list.
#'
#' @section Initialization:
#' Creating new objects of this class should be done via the `get_list` method of the [ms_site] class. Calling the `new()` method for this class only constructs the R object; it does not call the Microsoft Graph API to retrieve or create the actual item.
#'
#' @section List querying:
#' `list_items` supports the following arguments to customise results returned by the query.
#' - `filter`: A string giving a logical expression to filter the rows to return. Note that column names used in the expression must be prefixed with `fields/` to distinguish them from item metadata.
#' - `select`: A string containing comma-separated column names to include in the returned data frame. If not supplied, includes all columns.
#' - `all_metadata`: If TRUE, the returned data frame will contain extended metadata as separate columns, while the data fields will be in a nested data frame named `fields`.
#' - `pagesize`: The number of results to return for each call to the REST endpoint. You can try reducing this argument below the default of 5000 if you are experiencing timeouts.
#'
#' For more information, see [Use query parameters](https://docs.microsoft.com/en-us/graph/query-parameters?view=graph-rest-1.0) at the Graph API reference.
#'
#' @seealso
#' [sharepoint_site], [ms_site]
#'
#' [Microsoft Graph overview](https://docs.microsoft.com/en-us/graph/overview),
#' [SharePoint sites API reference](https://docs.microsoft.com/en-us/graph/api/resources/sharepoint?view=graph-rest-1.0)
#'
#' @examples
#' \dontrun{
#'
#' site <- sharepoint_site("https://mycompany.sharepoint.com/sites/my-site-name")
#' lst <- site$get_list("mylist")
#'
#' lst$get_column_info()
#'
#' lst$list_items()
#' lst$list_items(filter="startswith(fields/firstname, 'John')", select="firstname,lastname")
#'
#' }
#' @format An R6 object of class `ms_sharepoint_list`, inheriting from `ms_object`.
#' @export
ms_sharepoint_list <- R6::R6Class("ms_sharepoint_list", inherit=ms_object,

public=list(

    initialize=function(token, tenant=NULL, properties=NULL)
    {
        self$type <- "list"
        private$api_type <- "lists"
        super$initialize(token, tenant, properties)
    },

    list_items=function(filter=NULL, select=NULL, all_metadata=FALSE, pagesize=5000)
    {
        select <- if(is.null(select))
            "fields"
        else paste0("fields(select=", paste0(select, collapse=","), ")")
        options <- list(expand=select, `$filter`=filter, `$top`=pagesize)
        headers <- httr::add_headers(Prefer="HonorNonIndexedQueriesWarningMayFailRandomly")

        items <- self$do_operation("items", options=options, headers, simplify=TRUE)
        df <- private$get_paged_list(items, simplify=TRUE)
        if(!all_metadata)
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
        op <- sub("/$", "", file.path(
            "sites", self$properties$parentReference$siteId,
            "lists", self$properties$id,
            op
        ))
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
