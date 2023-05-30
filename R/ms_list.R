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
#' - `delete(confirm=TRUE)`: Delete this list. By default, ask for confirmation first.
#' - `update(...)`: Update the list's properties in Microsoft Graph.
#' - `do_operation(...)`: Carry out an arbitrary operation on the list.
#' - `sync_fields()`: Synchronise the R object with the list metadata in Microsoft Graph.
#' - `list_items(filter, select, all_metadata, as_data_frame, pagesize)`: Queries the list and returns items as a data frame. See 'List querying' below.
#' - `get_column_info()`: Return a data frame containing metadata on the columns (fields) in the list.
#' - `get_item(id)`: Get an individual list item.
#' - `create_item(...)`: Create a new list item, using the named arguments as fields.
#' - `update_item(id, ...)`: Update the _data_ fields in the given item, using the named arguments. To update the item's metadata, use `get_item()` to retrieve the item object, then call its `update()` method.
#' - `delete_item(confirm=TRUE)`: Delete a list item. By default, ask for confirmation first.
#' - `bulk_import(data)`: Imports a data frame into the list.
#'
#' @section Initialization:
#' Creating new objects of this class should be done via the `get_list` method of the [`ms_site`] class. Calling the `new()` method for this class only constructs the R object; it does not call the Microsoft Graph API to retrieve or create the actual item.
#'
#' @section List querying:
#' `list_items` supports the following arguments to customise results returned by the query.
#' - `filter`: A string giving an [OData expression](https://learn.microsoft.com/en-us/graph/query-parameters#filter-parameter) to filter the rows to return. Note that column names used in the expression must be prefixed with `fields/` to distinguish them from item metadata.
#' - `n`: The maximum number of (filtered) results to return. If this is NULL, the `ms_graph_pager` iterator object is returned instead to allow manual iteration over the results.
#' - `select`: A string containing comma-separated column names to include in the returned data frame. If not supplied, includes all columns.
#' - `all_metadata`: If TRUE, the returned data frame will contain extended metadata as separate columns, while the data fields will be in a nested data frame named `fields`. This is always set to FALSE if `n=NULL` or `as_data_frame=FALSE`.
#' - `as_data_frame`: If FALSE, return the result as a list of individual `ms_list_item` objects, rather than a data frame.
#' - `pagesize`: The number of results to return for each call to the REST endpoint. You can try reducing this argument below the default of 5000 if you are experiencing timeouts.
#'
#' Note that the Graph API currently doesn't support retrieving item attachments.
#'
#' @seealso
#' [`get_sharepoint_site`], [`ms_site`], [`ms_list_item`]
#'
#' [Microsoft Graph overview](https://learn.microsoft.com/en-us/graph/overview),
#' [SharePoint sites API reference](https://learn.microsoft.com/en-us/graph/api/resources/sharepoint?view=graph-rest-1.0)
#'
#' @examples
#' \dontrun{
#'
#' site <- get_sharepoint_site("My site")
#' lst <- site$get_list("mylist")
#'
#' lst$get_column_info()
#'
#' lst$list_items()
#' lst$list_items(filter="startswith(fields/firstname, 'John')", select="firstname,lastname")
#'
#' lst$create_item(firstname="Mary", lastname="Smith")
#' lst$get_item("item-id")
#' lst$update_item("item_id", firstname="Eliza")
#' lst$delete_item("item_id")
#'
#' df <- data.frame(
#'     firstname=c("Satya", "Mark", "Tim", "Jeff", "Sundar"),
#'     lastname=c("Nadella", "Zuckerberg", "Cook", "Bezos", "Pichai")
#' )
#' lst$bulk_import(df)
#'
#' }
#' @format An R6 object of class `ms_list`, inheriting from `ms_object`.
#' @export
ms_list <- R6::R6Class("ms_list", inherit=ms_object,

public=list(

    initialize=function(token, tenant=NULL, properties=NULL)
    {
        self$type <- "list"
        private$api_type <- file.path("sites", properties$parentReference$siteId, "lists")
        super$initialize(token, tenant, properties)
    },

    list_items=function(filter=NULL, select=NULL, all_metadata=FALSE, as_data_frame=TRUE, n=Inf, pagesize=5000)
    {
        select <- if(is.null(select))
            "fields"
        else paste0("fields(select=", paste0(select, collapse=","), ")")
        options <- list(expand=select, `$filter`=filter, `$top`=pagesize)
        headers <- httr::add_headers(Prefer="HonorNonIndexedQueriesWarningMayFailRandomly")

        pager <- self$get_list_pager(self$do_operation("items", options=options, headers, simplify=as_data_frame),
            site_id=self$properties$parentReference$siteId,
            list_id=self$properties$id)

        # get item list, or return the iterator immediately if n is NULL
        df <- extract_list_values(pager, n)
        if(is.null(n))
            return(df)

        if(as_data_frame && !all_metadata)
            df$fields
        else df
    },

    create_item=function(...)
    {
        fields <- list(...)
        res <- self$do_operation("items", body=list(fields=fields), http_verb="POST")
        invisible(ms_list_item$new(self$token, self$tenant, res,
            site_id=self$properties$parentReference$siteId,
            list_id=self$properties$id))
    },

    get_item=function(id)
    {
        res <- self$do_operation(file.path("items", id), options=list(expand="fields"))
        ms_list_item$new(self$token, self$tenant, res,
            site_id=self$properties$parentReference$siteId,
            list_id=self$properties$id)
    },

    update_item=function(id, ...)
    {
        fields <- list(...)
        self$get_item(id)$update(fields=list(...))
    },

    delete_item=function(id, confirm=TRUE)
    {
        self$get_item(id)$delete(confirm=confirm)
    },

    bulk_import=function(data)
    {
        stopifnot("Must supply a data frame"=is.data.frame(data))
        invisible(lapply(seq_len(nrow(data)), function(i) do.call(self$create_item, data[i, , drop=FALSE])))
    },

    get_column_info=function()
    {
        res <- self$do_operation(options=list(expand="columns"), simplify=TRUE)
        res$columns
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
