#' OneDrive clients
#'
#' @param tenant For `business_onedrive`, the name of your Azure Active Directory (AAD) tenant. If not supplied, use the default tenant for your currently logged-in account.
#' @details
#' These functions provide easy access to your OneDrive filesystem in the cloud. Use `personal_onedrive` to access the drive for your personal account, and `business_onedrive` to access the drive for your work or school account. For the latter, your organisation must have an appropriate Microsoft 365 license.
#'
#' The default "common" tenant for `business_onedrive` attempts to detect your actual tenant from your currently logged-in account. This may not always succeed, for example if you have a personal account that is also a guest account in a tenant. In this case, supply the actual tenant name.
#' @return
#' An object of class `ms_drive`.
#' @seealso
#' [ms_drive]
#' @examples
#' \dontrun{
#'
#' od <- personal_onedrive()
#' odb <- business_onedrive("myaadtenant")
#' odb$list_items()
#'
#' }
#' @rdname onedrive
#' @export
personal_onedrive <- function()
{
    login <- try(get_graph_login("consumers", refresh=FALSE), silent=TRUE)
    if(inherits(login, "try-error"))
        login <- create_graph_login("consumers", app=.azurer_graph_app_id)

    login$get_user()$get_drive()
}

#' @rdname onedrive
#' @export
business_onedrive <- function(tenant="common")
{
    login <- try(get_graph_login(tenant, refresh=FALSE), silent=TRUE)
    if(inherits(login, "try-error"))
        login <- create_graph_login(tenant, app=.az_cli_app_id)

    login$get_user()$get_drive()
}
