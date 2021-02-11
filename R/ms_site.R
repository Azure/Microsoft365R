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
#' - `list_drives()`: List the drives (shared document libraries) associated with this site.
#' - `get_drive(drive_id)`: Retrieve a shared document library for this site. If the ID is not specified, this returns the default document library.
#' - `list_subsites()`: List the subsites of this site.
#' - `get_lists()`: Returns the lists that are part of this site.
#' - `get_list(list_name, list_id)`: Returns a specific list, either by name or ID.
#' - `get_group()`: Retrieve the Microsoft 365 group associated with the site.
#'
#' @section Initialization:
#' Creating new objects of this class should be done via the `get_sharepoint_site` method of the [ms_graph] or [az_group] classes. Calling the `new()` method for this class only constructs the R object; it does not call the Microsoft Graph API to retrieve or create the actual site.
#'
#' @seealso
#' [ms_graph], [ms_drive], [az_user]
#'
#' [Microsoft Graph overview](https://docs.microsoft.com/en-us/graph/overview),
#' [SharePoint sites API reference](https://docs.microsoft.com/en-us/graph/api/resources/sharepoint?view=graph-rest-1.0)
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

    list_drives=function()
    {
        res <- private$get_paged_list(self$do_operation("drives"))
        private$init_list_objects(res, "drive")
    },

    get_drive=function(drive_id=NULL)
    {
        op <- if(is.null(drive_id))
            "drive"
        else file.path("drives", drive_id)
        ms_drive$new(self$token, self$tenant, self$do_operation(op))
    },

    list_subsites=function()
    {
        res <- private$get_paged_list(self$do_operation("sites"))
        private$init_list_objects(res, "site")
    },

    get_lists=function()
    {
        res <- private$get_paged_list(self$do_operation("lists"))
        private$init_list_objects(res, "list")
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
            stop("Unable to get group", call.=FALSE)
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
