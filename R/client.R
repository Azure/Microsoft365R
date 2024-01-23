#' Login clients for Microsoft 365
#'
#' Microsoft365R provides functions for logging into each Microsoft 365 service.
#'
#' @param tenant For `get_business_onedrive`, `get_sharepoint_site` and `get_team`, the name of your Azure Active Directory (AAD) tenant. If not supplied, use the value of the `CLIMICROSOFT365_TENANT` environment variable, or "common" if that is unset.
#' @param app A custom app registration ID to use for authentication. See below.
#' @param scopes The Microsoft Graph scopes (permissions) to obtain. It should never be necessary to change these.
#' @param site_name,site_url,site_id For `get_sharepoint_site`, either the name, web URL or ID of the SharePoint site to retrieve. Supply exactly one of these.
#' @param team_name,team_id For `get_team`, either the name or ID of the team to retrieve. Supply exactly one of these.
#' @param shared_mbox_id,shared_mbox_name,shared_mbox_email For `get_business_outlook`, an ID/principal name/email address. Supply exactly one of these to retrieve a shared mailbox. If all are NULL (the default), retrieve your own mailbox.
#' @param chat_id For `get_chat`, the ID of a group, one-on-one or meeting chat in Teams.
#' @param token An AAD OAuth token object, of class `AzureAuth::AzureToken`. If supplied, the `tenant`, `app`, `scopes` and `...` arguments will be ignored. See "Authenticating with a token" below.
#' @param ... Optional arguments that will ultimately be passed to [`AzureAuth::get_azure_token`].
#' @details
#' These functions provide easy access to the various collaboration services that are part of Microsoft 365. On first use, they will call your web browser to authenticate with Azure Active Directory, in a similar manner to other web apps. You will get a dialog box asking for permission to access your information. You only have to authenticate once; your credentials will be saved and reloaded in subsequent sessions.
#'
#' When authenticating, you can pass optional arguments in `...` which will ultimately be received by `AzureAuth::get_azure_token`. In particular, if your machine doesn't have a web browser available to authenticate with (for example if you are in a remote RStudio Server session), pass `auth_type="device_code"` which is intended for such scenarios.
#'
#' ## Authenticating to Microsoft 365 Business services
#' Authenticating to Microsoft 365 Business services (Teams, SharePoint and business OneDrive/Outlook) has some specific complexities.
#'
#' The default "common" tenant for `get_team`, `get_business_onedrive` and `get_sharepoint_site` attempts to detect your actual tenant from your saved credentials in your browser. This may not always succeed, for example if you have a personal account that is also a guest account in a tenant. In this case, supply the actual tenant name, either in the `tenant` argument or in the `CLIMICROSOFT365_TENANT` environment variable. The latter allows sharing authentication details with the [CLI for Microsoft 365](https://pnp.github.io/cli-microsoft365/).
#'
#' The default when authenticating to these services is for Microsoft365R to use its own internal app ID. As an alternative, you (or your admin) can create your own app registration in Azure: for use in a local session, it should have a native redirect URI of `http://localhost:1410`, and the "public client" option should be enabled if you want to use the device code authentication flow. You can supply your app ID either via the `app` argument, or in the environment variable `CLIMICROSOFT365_AADAPPID`.
#'
#' ## Authenticating with a token
#' In some circumstances, it may be desirable to carry out authentication/authorization as a separate step prior to  making requests to the Microsoft 365 REST API. This holds in a Shiny app, for example, since only the UI part can talk to the browser while the server part does the rest of the work. Another scenario is if the refresh token lifetime set by your org is too short, so that the token expires in between R sessions.
#'
#' In this case, you can authenticate by obtaining a new token with `AzureAuth::get_azure_token`, and passing the token object to the client function. Note that the token is accepted as-is; no checks are performed that it has the correct permissions for the service you're using.
#'
#' When calling `get_azure_token`, the scopes you should use are those given in the `scopes` argument for each client function, and the API host is `https://graph.microsoft.com/`. The Microsoft365R internal app ID is `d44a05d5-c6a5-4bbb-82d2-443123722380`, while that for the CLI for Microsoft 365 is `31359c7f-bd7e-475c-86db-fdb8c937548e`. However, these app IDs **only** work for a local R session; you must create your own app registration if you want to use the package inside a Shiny app.
#'
#' See the examples below, and also the vignette "Using Microsoft365R in a Shiny app" for a more detailed rundown on combining Microsoft365R and Shiny.
#'
#' ## Clearing the cache
#' Deleting your cached credentials is a way of rebooting the authentication process, if you are repeatedly encountering errors. To do this, call [`AzureAuth::clean_token_directory`], then try logging in again. You may also need to clear your browser's cookies, if you are authenticating interactively.
#'
#' @return
#' For `get_personal_onedrive` and `get_business_onedrive`, an R6 object of class `ms_drive`.
#'
#' For `get_sharepoint_site`, an R6 object of class `ms_site`; for `list_sharepoint_sites`, a list of such objects.
#'
#' For `get_team`, an R6 object of class `ms_team`; for `list_teams`, a list of such objects.
#' @seealso
#' [`ms_drive`], [`ms_site`], [`ms_team`], [`ms_chat`], [Microsoft365R global options][microsoft365r_options]
#'
#' [`add_methods`] for the associated methods that this package adds to the base AzureGraph classes.
#'
#' The "Authentication" vignette has more details on the authentication process, including troubleshooting and fixes for common problems. The "Using Microsoft365R in a Shiny app" vignette has further Shiny-specific information, including how to configure the necessary app registration in Azure Active Directory.
#'
#' [CLI for Microsoft 365](https://pnp.github.io/cli-microsoft365/) -- a commandline tool for managing Microsoft 365
#' @examples
#' \dontrun{
#'
#' get_personal_onedrive()
#'
#' # authenticating without a browser
#' get_personal_onedrive(auth_type="device_code")
#'
#' odb <- get_business_onedrive("mycompany")
#' odb$list_items()
#'
#' mysite <- get_sharepoint_site("My site", tenant="mycompany")
#' mysite <- get_sharepoint_site(site_url="https://mycompany.sharepoint.com/sites/my-site-url")
#' mysite$get_drive()$list_items()
#'
#' myteam <- get_team("My team", tenant="mycompany")
#' myteam$list_channels()
#' myteam$get_drive()$list_items()
#'
#' # retrieving chats
#' get_chat("chat-id")
#' list_chats()
#'
#' # you can also use your own app registration ID:
#' get_business_onedrive(app="app_id")
#' get_sharepoint_site("My site", app="app_id")
#'
#' # using the app ID for the CLI for Microsoft 365: set a global option
#' options(microsoft365r_use_cli_app_id=TRUE)
#' get_business_onedrive()
#' get_sharepoint_site("My site")
#' get_team("My team")
#'
#' # authenticating separately to working with the MS365 API
#' scopes <- c(
#'     "https://graph.microsoft.com/Files.ReadWrite.All",
#'     "https://graph.microsoft.com/User.Read",
#'     "openid", "offline_access"
#' )
#' app <- "d44a05d5-c6a5-4bbb-82d2-443123722380" # for local use only
#' token <- AzureAuth::get_azure_token(scopes, "mycompany", app, version=2)
#' get_business_onedrive(token=token)
#'
#' }
#' @rdname client
#' @export
get_personal_onedrive <- function(app=.microsoft365r_app_id,
                                  scopes=c("Files.ReadWrite.All", "User.Read"),
                                  token=NULL,
                                  ...)
{
    do_login("consumers", app, scopes, token, ...)$get_user()$get_drive()
}

#' @rdname client
#' @export
get_business_onedrive <- function(tenant=Sys.getenv("CLIMICROSOFT365_TENANT", "common"),
                                  app=Sys.getenv("CLIMICROSOFT365_AADAPPID"),
                                  scopes=c("Files.ReadWrite.All", "User.Read"),
                                  token=NULL,
                                  ...)
{
    app <- choose_app(app)
    scopes <- set_default_scopes(scopes, app)
    do_login(tenant, app, scopes, token, ...)$get_user()$get_drive()
}

#' @rdname client
#' @export
get_sharepoint_site <- function(site_name=NULL, site_url=NULL, site_id=NULL,
                                tenant=Sys.getenv("CLIMICROSOFT365_TENANT", "common"),
                                app=Sys.getenv("CLIMICROSOFT365_AADAPPID"),
                                scopes=c("Group.ReadWrite.All", "Directory.Read.All",
                                         "Sites.ReadWrite.All", "Sites.Manage.All"),
                                token=NULL,
                                ...)
{
    assert_one_arg(site_name, site_url, site_id, msg="Supply exactly one of site name, URL or ID")
    app <- choose_app(app)
    scopes <- set_default_scopes(scopes, app)
    login <- do_login(tenant, app, scopes, token, ...)

    if(!is.null(site_name))
    {
        filter <- sprintf("displayName eq '%s'", site_name)
        mysites <- login$get_user()$list_sharepoint_sites(filter=filter)
        if(length(mysites) == 0)
            stop("Site '", site_name, "' not found", call.=FALSE)
        else if(length(mysites) > 1)
            stop("Site name '", site_name, "' is not unique", call.=FALSE)
        mysites[[1]]
    }
    else login$get_sharepoint_site(site_url, site_id)
}

#' @rdname client
#' @export
list_sharepoint_sites <- function(tenant=Sys.getenv("CLIMICROSOFT365_TENANT", "common"),
                                  app=Sys.getenv("CLIMICROSOFT365_AADAPPID"),
                                  scopes=c("Group.ReadWrite.All", "Directory.Read.All",
                                           "Sites.ReadWrite.All", "Sites.Manage.All"),
                                  token=NULL,
                                  ...)
{
    app <- choose_app(app)
    scopes <- set_default_scopes(scopes, app)
    login <- do_login(tenant, app, scopes, token, ...)

    login$get_user()$list_sharepoint_sites()
}

#' @rdname client
#' @export
get_team <- function(team_name=NULL, team_id=NULL,
                     tenant=Sys.getenv("CLIMICROSOFT365_TENANT", "common"),
                     app=Sys.getenv("CLIMICROSOFT365_AADAPPID"),
                     scopes=c("Group.ReadWrite.All", "Directory.Read.All"),
                     token=NULL,
                     ...)
{
    assert_one_arg(team_name, team_id, msg="Supply exactly one of team name or ID")
    app <- choose_app(app)
    scopes <- set_default_scopes(scopes, app)
    login <- do_login(tenant, app, scopes, token, ...)

    if(!is.null(team_name))
    {
        filter <- sprintf("displayName eq '%s'", team_name)
        myteams <- login$get_user()$list_teams(filter=filter)
        # robustify against filter not working
        wch <- which(sapply(myteams, function(obj) obj$properties$displayName == team_name))
        if(length(wch) == 0)
            stop("Team '", team_name, "' not found", call.=FALSE)
        else if(length(wch) > 1)
            stop("Team name '", team_name, "' is not unique", call.=FALSE)
        myteams[[wch]]
    }
    else login$get_team(team_id)
}

#' @rdname client
#' @export
list_teams <- function(tenant=Sys.getenv("CLIMICROSOFT365_TENANT", "common"),
                       app=Sys.getenv("CLIMICROSOFT365_AADAPPID"),
                       scopes=c("Group.ReadWrite.All", "Directory.Read.All"),
                       token=NULL,
                       ...)
{
    app <- choose_app(app)
    scopes <- set_default_scopes(scopes, app)
    login <- do_login(tenant, app, scopes, token, ...)

    login$get_user()$list_teams()
}

#' @rdname client
#' @export
get_personal_outlook <- function(app=.microsoft365r_app_id,
                                 scopes=c("Mail.Send", "Mail.ReadWrite", "User.Read"),
                                 token=NULL,
                                 ...)
{
    do_login("consumers", app, scopes, token, ...)$get_user()$get_outlook()
}

#' @rdname client
#' @export
get_business_outlook <- function(tenant=Sys.getenv("CLIMICROSOFT365_TENANT", "common"),
                                 app=.microsoft365r_app_id,
                                 shared_mbox_id=NULL, shared_mbox_name=NULL, shared_mbox_email=NULL,
                                 scopes=c("User.Read", "Mail.Send", "Mail.ReadWrite"),
                                 token=NULL,
                                 ...)
{
    if(!is.null(shared_mbox_id) || !is.null(shared_mbox_name) || !is.null(shared_mbox_email))
        scopes <- c(scopes, "Mail.Send.Shared", "Mail.ReadWrite.Shared")

    do_login(tenant, app, scopes, token, ...)$
        get_user(user_id=shared_mbox_id, name=shared_mbox_name, email=shared_mbox_email)$
        get_outlook()
}

#' @rdname client
#' @export
get_chat <- function(chat_id,
                     tenant=Sys.getenv("CLIMICROSOFT365_TENANT", "common"),
                     app=.microsoft365r_app_id,
                     scopes=c("User.Read", "Directory.Read.All", "Chat.ReadWrite"),
                     token=NULL,
                     ...)
{
    do_login(tenant, app, scopes, token, ...)$get_user()$get_chat(chat_id)
}


#' @rdname client
#' @export
list_chats <- function(tenant=Sys.getenv("CLIMICROSOFT365_TENANT", "common"),
                       app=.microsoft365r_app_id,
                       scopes=c("User.Read", "Directory.Read.All", "Chat.ReadWrite"),
                       token=NULL,
                       ...)
{
    do_login(tenant, app, scopes, token, ...)$get_user()$list_chats()
}


do_login <- function(tenant, app, scopes, token, ...)
{
    # bypass AzureGraph login caching if token provided
    if(!is.null(token))
    {
        if(!AzureAuth::is_azure_token(token))
            stop("Invalid token object supplied")
        return(ms_graph$new(token=token))
    }

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


assert_one_arg <- function(..., msg=NULL)
{
    arglst <- list(...)
    nulls <- sapply(arglst, is.null)
    if(sum(!nulls) != 1)
        stop(msg, call.=FALSE)
}


set_default_scopes <- function(scopes, app)
{
    if(app %in% c(.cli_microsoft365_app_id, get(".az_cli_app_id", getNamespace("AzureGraph"))))
        scopes <- ".default"
    scopes
}
