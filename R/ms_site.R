#' Office 365 SharePoint site
#'
#' Class representing a SharePoint site.
#'
#' @docType class
#' @section Fields:
#' - `token`: The token used to authenticate with the Graph host.
#' - `tenant`: The Azure Active Directory tenant for this site.
#' - `type`: always "site" for a site object.
#' - `properties`: The site properties.
#' @section Methods:
#' - `new(...)`: Initialize a new site object. Do not call this directly; see 'Initialization' below.
#' - `delete(confirm=TRUE)`: Delete a site. By default, ask for confirmation first.
#' - `update(...)`: Update the site metadata in Microsoft Graph.
#' - `do_operation(...)`: Carry out an arbitrary operation on the site.
#' - `sync_fields()`: Synchronise the R object with the site metadata in Microsoft Graph.
#' - `list_drives(filter=NULL, n=Inf)`: List the drives (shared document libraries) associated with this site.
#' - `get_drive(drive_name, drive_id)`: Retrieve a shared document library for this site. If the name and ID are not specified, this returns the default document library.
#' - `list_subsites(filter=NULL, n=Inf)`: List the subsites of this site.
#' - `get_lists(filter=NULL, n=Inf)`: Returns the lists that are part of this site.
#' - `get_list(list_name, list_id)`: Returns a specific list, either by name or ID.
#' - `get_group()`: Retrieve the Microsoft 365 group associated with the site, if it exists. A site that backs a private Teams channel will not have a group associated with it.
#'
#' @section Initialization:
#' Creating new objects of this class should be done via the `get_sharepoint_site` method of the [`ms_graph`] or [`az_group`] classes. Calling the `new()` method for this class only constructs the R object; it does not call the Microsoft Graph API to retrieve or create the actual site.
#'
#' @section List methods:
#' All `list_*` methods have `filter` and `n` arguments to limit the number of results. The former should be an [OData expression](https://learn.microsoft.com/en-us/graph/query-parameters#filter-parameter) as a string to filter the result set on. The latter should be a number setting the maximum number of (filtered) results to return. The default values are `filter=NULL` and `n=Inf`. If `n=NULL`, the `ms_graph_pager` iterator object is returned instead to allow manual iteration over the results.
#'
#' Support in the underlying Graph API for OData queries is patchy. Not all endpoints that return lists of objects support filtering, and if they do, they may not allow all of the defined operators. If your filtering expression results in an error, you can carry out the operation without filtering and then filter the results on the client side.
#' @seealso
#' [`ms_graph`], [`ms_drive`], [`az_user`]
#'
#' [Microsoft Graph overview](https://learn.microsoft.com/en-us/graph/overview),
#' [SharePoint sites API reference](https://learn.microsoft.com/en-us/graph/api/resources/sharepoint?view=graph-rest-1.0)
#'
#' @examples
#' \dontrun{
#'
#' site <- get_sharepoint_site("My site")
#' site$list_drives()
#' site$get_drive()
#'
#' }
#' @format An R6 object of class `ms_site`, inheriting from `ms_object`.
#' @export
ms_site <- R6::R6Class("ms_site", inherit=ms_object,

public=list(

    initialize=function(token, tenant=NULL, properties=NULL)
    {
        self$type <- "site"
        private$api_type <- "sites"
        super$initialize(token, tenant, properties)
    },

    list_drives=function(filter=NULL, n=Inf)
    {
        private$make_basic_list("drives", filter, n)
    },

    get_drive=function(drive_name=NULL, drive_id=NULL)
    {
        if(!is.null(drive_name) && !is.null(drive_id))
            stop("Supply at most one of drive name or ID", call.=FALSE)
        if(!is.null(drive_name))
        {
            # filtering not yet supported for drives, do it in R
            drives <- self$list_drives()
            wch <- which(sapply(drives, function(drv) drv$properties$name == drive_name))
            if(length(wch) != 1)
                stop("Invalid drive name", call.=FALSE)
            return(drives[[wch]])
        }
        op <- if(is.null(drive_id))
            "drive"
        else file.path("drives", drive_id)
        ms_drive$new(self$token, self$tenant, self$do_operation(op))
    },

    list_subsites=function(filter=NULL, n=Inf)
    {
        private$make_basic_list("sites", filter, n)
    },

    get_lists=function(filter=NULL, n=Inf)
    {
        private$make_basic_list("lists", filter, n)
    },

    get_list=function(list_name=NULL, list_id=NULL)
    {
        assert_one_arg(list_name, list_id, msg="Supply exactly one of list name or ID")
        op <- if(!is.null(list_id))
            file.path("lists", list_id)
        else file.path("lists", curl::curl_escape(list_name))

        res <- self$do_operation(op)
        ms_list$new(self$token, self$tenant, res)
    },

    get_group=function()
    {
        filter <- sprintf("displayName eq '%s'", self$properties$displayName)
        res <- call_graph_endpoint(self$token, "groups", options=list(`$filter`=filter))$value
        if(length(res) != 1)
            stop("Unable to get Microsoft 365 group", call.=FALSE)
        az_group$new(self$token, self$tenant, res[[1]])
    },

    print=function(...)
    {
        cat("<Sharepoint site '", self$properties$displayName, "'>\n", sep="")
        cat("  directory id:", self$properties$id, "\n")
        cat("  web link:", self$properties$webUrl, "\n")
        cat("  description:", self$properties$description, "\n")
        cat("---\n")
        cat(format_public_methods(self))
        invisible(self)
    }
))
