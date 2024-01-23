#' SharePoint list item
#'
#' Class representing an item in a SharePoint list.
#'
#' @docType class
#' @section Fields:
#' - `token`: The token used to authenticate with the Graph host.
#' - `tenant`: The Azure Active Directory tenant for the parent drive.
#' - `type`: always "drive item" for a drive item object.
#' - `properties`: The item properties (data and metadata). This is a list; the item data can be found in the `fields` component.
#' @section Methods:
#' - `new(...)`: Initialize a new object. Do not call this directly; see 'Initialization' below.
#' - `delete(confirm=TRUE)`: Delete this item. By default, ask for confirmation first.
#' - `update(...)`: Update the item's properties (metadata) in Microsoft Graph. To update the list _data_, update the `fields` property. See the examples below.
#' - `do_operation(...)`: Carry out an arbitrary operation on the item.
#' - `sync_fields()`: Synchronise the R object with the item data and metadata in Microsoft Graph.
#'
#' @section Initialization:
#' Creating new objects of this class should be done via the `get_item` method of the [`ms_list`] class. Calling the `new()` method for this class only constructs the R object; it does not call the Microsoft Graph API to retrieve or create the actual item.
#'
#' @seealso
#' [`ms_graph`], [`ms_site`], [`ms_list`]
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
#' lst_items <- lst$list_items(as_data_frame=FALSE)
#'
#' item <- lst_items[[1]]
#'
#' item$update(fields=list(firstname="Mary"))
#'
#' # item data (plus some metadata mixed in)
#' item$properties$fields
#'
#' item$delete()
#'
#' }
#' @format An R6 object of class `ms_list_item`, inheriting from `ms_object`.
#' @export
ms_list_item <- R6::R6Class("ms_list_item", inherit=ms_object,

public=list(

    site_id=NULL,
    list_id=NULL,

    initialize=function(token, tenant=NULL, properties=NULL, site_id=NULL, list_id=NULL)
    {
        if(is.null(site_id) || is.null(list_id))
            stop("Must supply parent site and list IDs", call.=FALSE)
        self$type <- "list item"
        self$site_id <- site_id
        self$list_id <- list_id
        private$api_type <- file.path("sites", self$site_id, "lists", self$list_id, "items")
        super$initialize(token, tenant, properties)
    },

    print=function(...)
    {
        cat("<Sharepoint list item '", self$properties$fields$Title, "'>\n", sep="")
        cat("  directory id:", self$properties$id, "\n")
        cat("  web link:", self$properties$webUrl, "\n")
        cat("---\n")
        cat(format_public_methods(self))
        invisible(self)
    }
))

