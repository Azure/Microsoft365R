#' Deprecated client functions
#'
#' @param tenant For `business_onedrive` and `sharepoint_site`, the name of your Azure Active Directory (AAD) tenant. If not supplied, use the value of the `CLIMICROSOFT365_TENANT` environment variable, or "common" if that is unset.
#' @param app A custom app registration ID to use for authentication. For `personal_onedrive`, the default is to use Microsoft365R's internal app ID. For `business_onedrive` and `sharepoint_site`, see below.
#' @param scopes The Microsoft Graph scopes (permissions) to obtain.
#' @param site_url,site_id For `sharepoint_site`, the web URL and ID of the SharePoint site to retrieve. Supply one or the other, but not both.
#' @param ... Optional arguments to be passed to `AzureGraph::create_graph_login`.
#' @details
#' These functions have been replaced by [`get_personal_onedrive`], [`get_business_onedrive`] and [`get_sharepoint_site`]. They will be removed in a later version of the package.
#' @rdname Microsoft365R-deprecated
#' @aliases client-deprecated
#' @export
personal_onedrive <- function(app=.microsoft365r_app_id,
                              scopes=c("Files.ReadWrite.All", "User.Read"),
                              ...)
{
    .Deprecated("get_personal_onedrive")
    get_personal_onedrive(app=app, scopes=scopes, ...)
}


#' @rdname Microsoft365R-deprecated
#' @export
business_onedrive <- function(tenant=Sys.getenv("CLIMICROSOFT365_TENANT", "common"),
                              app=Sys.getenv("CLIMICROSOFT365_AADAPPID"),
                              scopes=".default",
                              ...)
{
    .Deprecated("get_business_onedrive")
    get_business_onedrive(tenant=tenant, app=app, scopes=scopes, ...)
}


#' @rdname Microsoft365R-deprecated
#' @export
sharepoint_site <- function(site_url=NULL, site_id=NULL,
                            tenant=Sys.getenv("CLIMICROSOFT365_TENANT", "common"),
                            app=Sys.getenv("CLIMICROSOFT365_AADAPPID"),
                            scopes=".default",
                            ...)
{
    .Deprecated("get_sharepoint_site")
    get_sharepoint_site(site_url=site_url, site_id=site_id, tenant=tenant, app=app, scopes=scopes, ...)
}
