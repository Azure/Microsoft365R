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
#' - `open()`: Open the item in your browser.
#' - `download(dest, overwrite)`: Download the file. Not applicable for a folder.
#' - `create_share_link(type, expiry, password, scope)`: Create a shareable link to the file or folder. See 'Sharing' below.
#'
#' @section Initialization:
#' Creating new objects of this class should be done via the `get_item_properties` method of the [ms_drive] class. Calling the `new()` method for this class only constructs the R object; it does not call the Microsoft Graph API to retrieve or create the actual item.
#'
#' @section Sharing:
#' `create_share_link(type, expiry, password, scope)` returns a shareable link to the item. Its arguments are
#' - `type`: Either "view" for a read-only link, "edit" for a read-write link, or "embed" for a link that can be embedded in a web page. The last one is only available for personal OneDrive.
#' - `expiry`: How long the link is valid for. The default is 7 days; you can set an alternative like "15 minutes", "24 hours", "2 weeks", "3 months", etc. To leave out the expiry date, set this to NULL.
#' - `scope`: Optionally the scope of the link, either "anonymous" or "organization". The latter allows only users in your AAD tenant to access the link, and is only available for OneDrive for Business or SharePoint.
#'
#' This function returns a URL to access the item, for `type="view"` or "`type=edit"`. For `type="embed"`, it returns a list with components `webUrl` containing the URL, and `webHtml` containing a HTML fragment to embed the link in an IFRAME.
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
#' # download the file to the working directory
#' myfile$download()
#'
#' # shareable links
#' myfile$create_share_link()
#' myfile$create_share_link(type="edit", expiry="24 hours")
#' myfile$create_share_link(password="Use-strong-passwords!")
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

    create_share_link=function(type=c("view", "edit", "embed"), expiry="7 days", password=NULL, scope=NULL)
    {
        body <- list(type=match.arg(type))
        if(!is.null(expiry))
        {
            expdate <- seq(Sys.time(), by=expiry, len=2)[2]
            expirationDateTime <- strftime(expdate, "%Y-%m-%dT%H:%M:%SZ", tz="GMT")
        }
        if(!is.null(password))
            body$password <- password
        if(!is.null(scope))
            body$scope <- scope
        res <- self$do_operation("createLink", body=body, http_verb="POST")
        if(type == "embed")
        {
            res$link$type <- NULL
            res$link
        }
        else res$link$webUrl
    },

    download=function(dest=self$properties$name, overwrite=FALSE)
    {
        filepath <- file.path(self$parentReference$path, self$properties$name)
        res <- self$do_operation("content", config=httr::write_disk(dest, overwrite=overwrite),
                                 http_status_handler="pass")
        if(httr::status_code(res) >= 300)
        {
            on.exit(file.remove(dest))
            httr::stop_for_status(res, paste0("complete operation. Message:\n",
                sub("\\.$", "", error_message(httr::content(res)))))
        }
        invisible(NULL)
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
