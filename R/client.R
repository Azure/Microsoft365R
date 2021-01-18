#' OneDrive and Sharepoint Online clients
#'
#' @param tenant For `business_onedrive` and `sharepoint_site`, the name of your Azure Active Directory (AAD) tenant. If not supplied, use the default tenant for your currently logged-in account.
#' @param app Optionally, a custom app registration ID to use for authentication.
#' @param site_url,site_id For `sharepoint_site`, the web URL and ID of the SharePoint site to retrieve. Supply one or the other, but not both.
#' @param ... Optional arguments to be passed to `AzureGraph::create_graph_login`.
#' @details
#' These functions provide easy access to OneDrive and SharePoint in the cloud. They work by loading your existing Microsoft Graph login, and if that isn't found, creating a new one using any arguments passed in `...`. In particular, if your machine doesn't have a web browser available to authenticate with (for example if you are in a remote RStudio Server session), pass `auth_type="device_code"` which is intended for such scenarios.
#'
#' Use `personal_onedrive` to access the drive for your personal account, and `business_onedrive` to access your OneDrive for Business. On first use, you will get a dialog box in your browser asking permission for Microsoft365R to access your information.
#'
#' The default "common" tenant for `business_onedrive` and `sharepoint_site` attempts to detect your actual tenant from your saved credentials in your browser. This may not always succeed, for example if you have a personal account that is also a guest account in a tenant. In this case, supply the actual tenant name.
#'
#' Note that OneDrive for Business and SharePoint are part of Microsoft 365 Business, so trying to access these services will fail if you don't have an appropriate license.
#' @return
#' An object of class `ms_drive`.
#' @seealso
#' [ms_drive], [AzureGraph::create_graph_login]
#' @examples
#' \dontrun{
#'
#' personal_onedrive()
#'
#' # authenticating without a browser
#' personal_onedrive(auth_type="device_code")
#'
#' odb <- business_onedrive("mycompany")
#' odb$list_items()
#'
#' site <- sharepoint_site("https://mycompany.sharepoint.com/sites/my-site-name", tenant="mycompany")
#' site$get_drive()$list_items()
#'
#' }
#' @rdname client
#' @export
personal_onedrive <- function(app=NULL, ...)
{
    if(is.null(app))
        app <- .microsoft365r_app_id
    scopes <- c("Files.ReadWrite.All", "User.Read")
    login <- try(get_graph_login("consumers", app=.microsoft365r_app_id, scopes=scopes, refresh=FALSE),
                 silent=TRUE)
    if(inherits(login, "try-error"))
        login <- create_graph_login("consumers", app=.microsoft365r_app_id, scopes=scopes, ...)

    login$get_user()$get_drive()
}

#' @rdname client
#' @export
business_onedrive <- function(tenant="common", app=NULL, ...)
{
    if(is.null(app))
        app <- .microsoft365r_app_id
    scopes <- c("Files.ReadWrite.All", "User.Read")
    login <- try(get_graph_login(tenant, app=app, scopes=scopes, refresh=FALSE), silent=TRUE)
    if(inherits(login, "try-error"))
        login <- create_graph_login(tenant, app=app, scopes=scopes, ...)

    login$get_user()$get_drive()
}

#' @rdname client
#' @export
sharepoint_site <- function(site_url=NULL, site_id=NULL, tenant="common", app=NULL, ...)
{
    if(is.null(app))
        app <- .az_cli_app_id
    login <- try(get_graph_login(tenant, app=app, scopes=".default", refresh=FALSE), silent=TRUE)
    if(inherits(login, "try-error"))
        login <- create_graph_login(tenant, app=app, scopes=".default", ...)

    login$get_sharepoint_site(site_url, site_id)
}
