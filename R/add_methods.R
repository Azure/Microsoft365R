# documentation separate from code because Roxygen can't handle adding methods to another package's R6 classes

#' Get SharePoint site or OneDrive
#'
#' Methods for the [AzureGraph::ms_graph], [AzureGraph::az_user] and [AzureGraph::az_group] classes.
#'
#' @rdname get_sharepoint_site
#' @name get_sharepoint_site
#' @aliases get_drive list_drives
#' @section Usage:
#' ```
#' ## R6 method for class 'ms_graph'
#' get_drive(drive_id)
#'
#' ## R6 method for class 'az_user'
#' get_drive(drive_id = NULL)
#'
#' ## R6 method for class 'az_group'
#' get_drive(drive_id = NULL)
#'
#' ## R6 method for class 'ms_graph'
#' get_sharepoint_site(site_url = NULL, site_id = NULL)
#'
#' ## R6 method for class 'az_group'
#' get_sharepoint_site()
#'
#' ## R6 method for class 'az_user'
#' list_drives()
#'
#' ## R6 method for class 'az_group'
#' list_drives()
#' ```
#' @section Arguments:
#' - `drive_id`: For `get_drive`, the ID of the drive or shared document library. For the `az_user` and `az_group` methods, if this is NULL the default drive/document library is returned.
#' - `site_url`,`site_id`: For `ms_graph$get_sharepoint_site()`, the URL and ID of the site. Provide one or the other, but not both.
#' @section Details:
#' `get_sharepoint_site` retrieves a SharePoint site object. The method for the top-level Graph client object requires that you provide either the site URL or ID. The method for an `az_group` object will retrieve the site associated with that group, if applicable.
#'
#' `get_drive` retrieves a OneDrive or shared document library, and `list_drives` retrieves all such drives/libraries that the user or group has access to. Whether these are personal or business drives depends on the tenant that was specified in `AzureGraph::get_graph_login()`/`create_graph_login()`: if this was "consumers" or "9188040d-6c67-4c5b-b112-36a304b66dad" (the equivalent GUID), it will be the personal OneDrive. See the examples below.
#'
#' Note that OneDrive for Business and SharePoint require a Microsoft 365 Business license, and are available for organisational tenants only.
#'
#' @section Value:
#' For `get_sharepoint_site`, an object of class `ms_site`. For `get_drive`, an object of class `ms_drive`. For `list_drives`, a list of `ms_drive` objects.
#' @seealso
#' [ms_site], [ms_drive], [az_user], [az_group]
#' @examples
#' \dontrun{
#'
#' # 'consumers' tenant -> personal OneDrive for a user
#' gr <- AzureGraph::get_graph_login("consumers", app="myapp")
#' me <- gr$get_user()
#' me$get_drive()
#'
#' # organisational tenant -> business OneDrive for a user
#' gr2 <- AzureGraph::get_graph_login("mycompany", app="myapp")
#' myuser <- gr2$get_user("username@mycompany.onmicrosoft.com")
#' myuser$get_drive()
#'
#' # get a site/drive directly from a URL/ID
#' gr2$get_sharepoint_site("https://mycompany.sharepoint.com/sites/my-site-name")
#' gr2$get_drive("drive-id")
#'
#' # site/drive(s) for a group
#' grp <- gr2$get_group("group-id")
#' grp$get_sharepoint_site()
#' grp$list_drives()
#' grp$get_drive()
#'
#' }
NULL

add_methods <- function()
{
    ms_graph$set("public", "get_sharepoint_site", overwrite=TRUE,
    function(site_url=NULL, site_id=NULL)
    {
        op <- if(is.null(site_url) && !is.null(site_id))
            file.path("sites", site_id)
        else if(!is.null(site_url) && is.null(site_id))
        {
            site_url <- httr::parse_url(site_url)
            file.path("sites", paste0(site_url$hostname, ":"), site_url$path)
        }
        else stop("Must supply either site ID or URL")

        ms_site$new(self$token, self$tenant, self$call_graph_endpoint(op))
    })

    ms_graph$set("public", "get_drive", overwrite=TRUE,
    function(drive_id)
    {
        op <- file.path("drives", drive_id)
        ms_drive$new(self$token, self$tenant, self$call_graph_endpoint(op))
    })

    az_user$set("public", "list_drives", overwrite=TRUE,
    function()
    {
        res <- private$get_paged_list(self$do_operation("drives"))
        private$init_list_objects(res, "drive")
    })

    az_user$set("public", "get_drive", overwrite=TRUE,
    function(drive_id=NULL)
    {
        op <- if(is.null(drive_id))
            "drive"
        else file.path("drives", drive_id)
        ms_drive$new(self$token, self$tenant, self$do_operation(op))
    })

    az_group$set("public", "get_sharepoint_site", overwrite=TRUE,
    function()
    {
        res <- self$do_operation("sites/root")
        ms_site$new(self$token, self$tenant, res)
    })

    az_group$set("public", "list_drives", overwrite=TRUE,
    function()
    {
        res <- private$get_paged_list(self$do_operation("drives"))
        private$init_list_objects(res, "drive")
    })

    az_group$set("public", "get_drive", overwrite=TRUE,
    function(drive_id=NULL)
    {
        op <- if(is.null(drive_id))
            "drive"
        else file.path("drives", drive_id)
        ms_drive$new(self$token, self$tenant, self$do_operation(op))
    })
}
