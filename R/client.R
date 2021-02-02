#' OneDrive and Sharepoint Online clients
#'
#' @param tenant For `business_onedrive` and `sharepoint_site`, the name of your Azure Active Directory (AAD) tenant. If not supplied, use the value of the `CLIMICROSOFT365_TENANT` environment variable, or "common" if that is unset.
#' @param app A custom app registration ID to use for authentication. For `personal_onedrive`, the default is to use Microsoft365R's internal app ID. For `business_onedrive` and `sharepoint_site`, see below.
#' @param scopes The Microsoft Graph scopes (permissions) to obtain.
#' @param site_url,site_id For `sharepoint_site`, the web URL and ID of the SharePoint site to retrieve. Supply one or the other, but not both.
#' @param team_name,team_id For `team`, the name and ID of the team to retrieve. Supply one or the other, but not both.
#' @param ... Optional arguments to be passed to `AzureGraph::create_graph_login`.
#' @details
#' `personal_onedrive`, `business_onedrive` and `sharepoint_site` provide easy access to OneDrive, OneDrive for Business, and SharePoint Online respectively. On first use, they will call your web browser to authenticate with Azure Active Directory, in a similar manner to other web apps. You will get a dialog box asking for permission to access your information. You only have to authenticate once per client; your credentials will be saved and reloaded in subsequent sessions.
#'
#' When authenticating, you can pass optional arguments in `...` which will ultimately be received by `AzureAuth::get_azure_token`. In particular, if your machine doesn't have a web browser available to authenticate with (for example if you are in a remote RStudio Server session), pass `auth_type="device_code"` which is intended for such scenarios.
#'
#' @section Authenticating to Microsoft 365 Business services:
#' Authenticating to Microsoft 365 Business services (SharePoint and OneDrive for Business) has some specific complexities.
#'
#' The default "common" tenant for `business_onedrive` and `sharepoint_site` attempts to detect your actual tenant from your saved credentials in your browser. This may not always succeed, for example if you have a personal account that is also a guest account in a tenant. In this case, supply the actual tenant name, either in the `tenant` argument or in the `CLIMICROSOFT365_TENANT` environment variable. The latter allows sharing authentication details with the [CLI for Microsoft 365](https://pnp.github.io/cli-microsoft365/).
#'
#' The default when authenticating to these services is for Microsoft365R to use its own internal app ID. Depending on your organisation's security policy, you may have to get an admin to grant it access to your tenant. As an alternative to the default app ID, you (or your admin) can create your own app registration: it should have a native redirect URI of `http://localhost:1410`, and the "public client" option should be enabled if you want to use the device code authentication flow. You can supply your app ID either via the `app` argument, or in the environment variable `CLIMICROSOFT365_AADAPPID`.
#'
#' If creating your own app registration is impractical, it's possible to work around access issues by piggybacking on the CLI for Microsoft365. By setting the R option `microsoft365r_use_cli_app_id` to a non-NULL value, authentication will be done using the CLI's app ID. Technically this app still requires admin approval, but it is in widespread use and so may already be allowed in your organisation. Be warned that this solution may draw the attention of your admin!
#'
#' @return
#' For `personal_onedrive` and `business_onedrive`, an object of class `ms_drive`. For `sharepoint_site`, an object of class `ms_site`.
#' @seealso
#' [ms_drive], [ms_site], [AzureGraph::create_graph_login], [AzureAuth::get_azure_token]
#'
#' [CLI for Microsoft 365](https://pnp.github.io/cli-microsoft365/) -- a commandline tool for managing Microsoft 365
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
#' # using the app ID for the CLI for Microsoft 365: set a global option
#' options(microsoft365r_use_cli_app_id=TRUE)
#' business_onedrive()
#' sharepoint_site("https://mycompany.sharepoint.com/sites/my-site-name")
#'
#' }
#' @rdname client
#' @export
personal_onedrive <- function(app=.microsoft365r_app_id,
                              scopes=c("Files.ReadWrite.All", "User.Read"),
                              ...)
{
    do_login("consumers", app, scopes, ...)$get_user()$get_drive()
}

#' @rdname client
#' @export
business_onedrive <- function(tenant=Sys.getenv("CLIMICROSOFT365_TENANT", "common"),
                              app=Sys.getenv("CLIMICROSOFT365_AADAPPID"),
                              scopes=".default",
                              ...)
{
    app <- choose_app(app)
    do_login(tenant, app, scopes, ...)$get_user()$get_drive()
}

#' @rdname client
#' @export
sharepoint_site <- function(site_url=NULL, site_id=NULL,
                            tenant=Sys.getenv("CLIMICROSOFT365_TENANT", "common"),
                            app=Sys.getenv("CLIMICROSOFT365_AADAPPID"),
                            scopes=".default",
                            ...)
{
    app <- choose_app(app)
    do_login(tenant, app, scopes, ...)$get_sharepoint_site(site_url, site_id)
}

#' @rdname client
#' @export
team <- function(team_name=NULL, team_id=NULL,
                 tenant=Sys.getenv("CLIMICROSOFT365_TENANT", "common"),
                 app=Sys.getenv("CLIMICROSOFT365_AADAPPID"),
                 scopes=".default",
                 ...)
{
    app <- choose_app(app)
    login <- do_login(tenant, app, scopes, ...)

    if(!is.null(team_name) && is.null(team_id))
        login$get_user()$get_team(team_name)
    else if(is.null(team_name) && !is.null(team_id))
        login$get_team(team_id)
    else stop("Must supply either team name or ID", call.=FALSE)
}


do_login <- function(tenant, app, scopes, ...)
{
    login <- try(get_graph_login(tenant, app=app, scopes=scopes, refresh=FALSE), silent=TRUE)
    if(inherits(login, "try-error"))
        login <- create_graph_login(tenant, app=app, scopes=scopes, ...)
    login
}


choose_app <- function(app)
{
    if(is.null(app) || app == "")
    {
        if(!is.null(getOption("microsoft365r_use_cli_app_id")))
            .cli_microsoft365_app_id
        else .microsoft365r_app_id
    }
    else app
}
