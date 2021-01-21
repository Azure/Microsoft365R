#' OneDrive and Sharepoint Online clients
#'
#' @param tenant For `business_onedrive` and `sharepoint_site`, the name of your Azure Active Directory (AAD) tenant. If not supplied, use the value of the `CLIMICROSOFT365_TENANT` environment variable, or "common" if that is unset.
#' @param app A custom app registration ID to use for authentication. If not supplied, use the value of the `CLIMICROSOFT365_AADAPPID` environment variable, or an internal app ID if that is unset.
#' @param scopes The Microsoft Graph scopes (permissions) to obtain.
#' @param site_url,site_id For `sharepoint_site`, the web URL and ID of the SharePoint site to retrieve. Supply one or the other, but not both.
#' @param ... Optional arguments to be passed to `AzureGraph::create_graph_login`.
#' @details
#' `personal_onedrive`, `business_onedrive` and `sharepoint_site` provide easy access to OneDrive, OneDrive for Business, and SharePoint Online respectively. On first use, they will call your web browser to authenticate with Azure Active Directory, in a similar manner to other web apps. You will get a dialog box asking for permission to access your information. You only have to authenticate once per client; your credentials will be saved and reloaded in subsequent sessions.
#'
#' When authenticating, you can pass optional arguments in `...` which will ultimately be received by `AzureAuth::get_azure_token`. In particular, if your machine doesn't have a web browser available to authenticate with (for example if you are in a remote RStudio Server session), pass `auth_type="device_code"` which is intended for such scenarios.
#'
#' The default "common" tenant for `business_onedrive` and `sharepoint_site` attempts to detect your actual tenant from your saved credentials in your browser. This may not always succeed, for example if you have a personal account that is also a guest account in a tenant. In this case, supply the actual tenant name, either in the `tenant` argument or in the `CLIMICROSOFT365_TENANT` environment variable. The latter allows sharing authentication details with the CLI for Microsoft 365.
#'
#' For authentication purposes, Microsoft365R is registered as an app in the "aicatr" AAD tenant. Depending on your organisation's security policy, you may have to get an admin to grant it access to your tenant.
#'
#' As an alternative to using the default app ID, you (or your admin) can create your own app registration. It should have a native redirect URI of `http://localhost:1410`, and the "public client" option should be enabled if you want to use the device code authentication flow. You can supply your app ID either via the `app` argument, or in the environment variable `CLIMICROSOFT365_AADAPPID`.
#'
#' In addition, for SharePoint (only) it's possible to use the Azure CLI app ID to access document libraries and lists. See the examples below. Be warned, however, that this may attract the attention of your admin!
#'
#' @return
#' An object of class `ms_drive`.
#' @seealso
#' [ms_drive], [AzureGraph::create_graph_login]
#'
#' [CLI for Microsoft 365](https://pnp.github.io/cli-microsoft365/) -- where the `CLIMICROSOFT365_AADAPPID` and `CLIMICROSOFT365_TENANT` environment variables come from
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
#'
#' # you can also use your own app registration ID:
#' business_onedrive(app="app_id")
#' sharepoint_site("https://mycompany.sharepoint.com/sites/my-site-name", app="app_id")
#'
#' # for SharePoint, a fallback is to use the Azure CLI app ID and the '.default' scope:
#' sharepoint_site("https://mycompany.sharepoint.com/sites/my-site-name",
#'     app=AzureGraph:::.az_cli_app_id,
#'     scopes=".default")
#'
#' }
#' @rdname client
#' @export
personal_onedrive <- function(app=Sys.getenv("CLIMICROSOFT365_AADAPPID", .microsoft365r_app_id),
                              scopes=c("Files.ReadWrite.All", "User.Read"),
                              ...)
{
    do_login("consumers", app, scopes, ...)$get_user()$get_drive()
}

#' @rdname client
#' @export
business_onedrive <- function(tenant=Sys.getenv("CLIMICROSOFT365_TENANT", "common"),
                              app=Sys.getenv("CLIMICROSOFT365_AADAPPID", .microsoft365r_app_id),
                              scopes=c("Files.ReadWrite.All", "User.Read"),
                              ...)
{
    do_login(tenant, app, scopes, ...)$get_user()$get_drive()
}

#' @rdname client
#' @export
sharepoint_site <- function(site_url=NULL, site_id=NULL,
                            tenant=Sys.getenv("CLIMICROSOFT365_TENANT", "common"),
                            app=Sys.getenv("CLIMICROSOFT365_AADAPPID", .microsoft365r_app_id),
                            scopes=c("Sites.ReadWrite.All", "User.Read"),
                            ...)
{
    do_login(tenant, app, scopes, ...)$get_sharepoint_site(site_url, site_id)
}


do_login <- function(tenant, app, scopes, ...)
{
    login <- try(get_graph_login(tenant, app=app, scopes=scopes, refresh=FALSE), silent=TRUE)
    if(inherits(login, "try-error"))
        login <- create_graph_login(tenant, app=app, scopes=scopes, ...)
    login
}
