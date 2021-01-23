#' SharePoint list item
#'
#' Class representing an item in a SharePoint list.
#'
#' @docType class
#' @section Fields:
#' - `token`: The token used to authenticate with the Graph host.
#' - `tenant`: The Azure Active Directory tenant for the parent drive.
#' - `type`: always "drive item" for a drive item object.
#' - `properties`: The item properties (metadata).
#' @section Methods:
#' - `new(...)`: Initialize a new object. Do not call this directly; see 'Initialization' below.
#' - `delete(confirm=TRUE)`: Delete this item. By default, ask for confirmation first.
#' - `update(...)`: Update the item's properties (metadata) in Microsoft Graph. To update the list _data_, update the `fields` property. See the examples below.
#' - `do_operation(...)`: Carry out an arbitrary operation on the item.
#' - `sync_fields()`: Synchronise the R object with the item metadata in Microsoft Graph.
#'
#' @section Initialization:
#' Creating new objects of this class should be done via the `get_item` method of the [ms_sharepoint_list] class. Calling the `new()` method for this class only constructs the R object; it does not call the Microsoft Graph API to retrieve or create the actual item.
#'
#' @seealso
#' [ms_graph], [ms_site], [ms_sharepoint_list]
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
#' lst_items <- lst$list_items(as_data_frame=FALSE)
#'
#' item <- lst_items[[1]]
#'
#' item$update(fields=list(firstname="Mary"))
#'
#' item$delete()
#'
#' }
#' @format An R6 object of class `ms_list_item`, inheriting from `ms_object`.
#' @export
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
