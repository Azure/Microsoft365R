#' File or folder in a drive
#'
#' Class representing an item (file or folder) in a OneDrive or SharePoint document library.
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
#' - `update(...)`: Update the item's properties (metadata) in Microsoft Graph.
#' - `do_operation(...)`: Carry out an arbitrary operation on the item.
#' - `sync_fields()`: Synchronise the R object with the item metadata in Microsoft Graph.
#' - `open()`: Opens the item in your browser.
#'
#' @section Initialization:
#' Creating new objects of this class should be done via the `get_item_properties` method of the [ms_drive] class. Calling the `new()` method for this class only constructs the R object; it does not call the Microsoft Graph API to retrieve or create the actual item.
#'
#' @seealso
#' [ms_graph], [ms_site], [ms_drive]
#'
#' [Microsoft Graph overview](https://docs.microsoft.com/en-us/graph/overview),
#' [OneDrive API reference](https://docs.microsoft.com/en-us/graph/api/resources/onedrive?view=graph-rest-1.0)
#'
#' @examples
#' \dontrun{
#'
#' # personal OneDrive
#' gr2 <- get_graph_login("consumers")
#' me <- gr2$get_user()
#' mydrv <- me$get_drive()
#'
#' myfile <- drv$get_item_properties("myfile.docx")
#' myfile$properties
#'
#' # rename a file
#' myfile$update(name="newname.docx")
#'
#' # open the file in the browser
#' myfile$open()
#'
#' # delete the file (will ask for confirmation first)
#' myfile$delete()
#'
#'
#' }
#' @format An R6 object of class `ms_drive_item`, inheriting from `ms_object`.
#' @export
ms_drive_item <- R6::R6Class("ms_drive_item", inherit=ms_object,

public=list(

    initialize=function(token, tenant=NULL, properties=NULL)
    {
        self$type <- "drive item"
        private$api_type <- file.path("drives", properties$parentReference$driveId, "items")
        super$initialize(token, tenant, properties)
    },

    open=function()
    {
        httr::BROWSE(self$properties$webUrl)
    },

    print=function(...)
    {
        file_or_dir <- if(!is.null(self$properties$folder)) "file folder" else "file"
        cat("<Drive item '", self$properties$name, "'>\n", sep="")
        cat("  directory id:", self$properties$id, "\n")
        cat("  web link:", self$properties$webUrl, "\n")
        cat("  type:", file_or_dir, "\n")
        cat("---\n")
        cat(format_public_methods(self))
        invisible(self)
    }
))
