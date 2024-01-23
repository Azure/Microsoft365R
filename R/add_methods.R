# documentation separate from code because Roxygen can't handle adding methods to another package's R6 classes

#' Microsoft 365 object accessor methods
#'
#' Methods for the [`AzureGraph::ms_graph`], [`AzureGraph::az_user`] and [`AzureGraph::az_group`] classes.
#'
#' @rdname add_methods
#' @name add_methods
#' @section Usage:
#' ```
#' ## R6 method for class 'az_user'
#' get_chat(chat_id)
#'
#' ## R6 method for class 'ms_graph'
#' get_drive(drive_id)
#'
#' ## R6 method for class 'az_user'
#' get_drive(drive_id = NULL)
#'
#' ## R6 method for class 'az_group'
#' get_drive(drive_name = NULL, drive_id = NULL)
#'
#' ## R6 method for class 'az_group'
#' get_plan(plan_title = NULL, plan_id = NULL)
#'
#' ## R6 method for class 'ms_graph'
#' get_sharepoint_site(site_url = NULL, site_id = NULL)
#'
#' ## R6 method for class 'az_group'
#' get_sharepoint_site()
#'
#' ## R6 method for class 'ms_graph'
#' get_team(team_id = NULL)
#'
#' ## R6 method for class 'az_group'
#' get_team()
#'
#' ## R6 method for class 'az_user'
#' list_chats(filter = NULL, n = Inf)
#'
#' ## R6 method for class 'az_user'
#' list_drives(filter = NULL, n = Inf)
#'
#' ## R6 method for class 'az_group'
#' list_drives(filter = NULL, n = Inf)
#'
#' ## R6 method for class 'az_group'
#' list_plans(filter = NULL, n = Inf)
#'
#' ## R6 method for class 'az_user'
#' list_sharepoint_sites(filter = NULL, n = Inf)
#'
#' ## R6 method for class 'az_user'
#' list_teams(filter = NULL, n = Inf)
#' ```
#' @section Arguments:
#' - `drive_name`,`drive_id`: For `get_drive`, the name or ID of the drive or shared document library. Note that only the `az_group` method  has the `drive_name` argument, as user drives do not have individual names (and most users will only have one drive anyway). For the `az_user` and `az_group` methods, leaving the argument(s) blank will return the default drive/document library.
#' - `site_url`,`site_id`: For `ms_graph$get_sharepoint_site()`, the URL and ID of the site. Provide one or the other, but not both.
#' - `team_name`,`team_id`: For `az_user$get_team()`, the name and ID of the site. Provide one or the other, but not both. For `ms_graph$get_team`, you must provide the team ID.
#' - `plan_title`,`plan_id`: For `az_group$get_plan()`, the title and ID of the site. Provide one or the other, but not both.
#' - `filter, n`: See 'List methods' below.
#' @section Details:
#' `get_sharepoint_site` retrieves a SharePoint site object. The method for the top-level Graph client class requires that you provide either the site URL or ID. The method for the `az_group` class will retrieve the site associated with that group, if applicable.
#'
#' `get_drive` retrieves a OneDrive or shared document library, and `list_drives` retrieves all such drives/libraries that the user or group has access to. Whether these are personal or business drives depends on the tenant that was specified in `AzureGraph::get_graph_login()`/`create_graph_login()`: if this was "consumers" or "9188040d-6c67-4c5b-b112-36a304b66dad" (the equivalent GUID), it will be the personal OneDrive. See the examples below.
#'
#' `get_plan` retrieves a plan (not to be confused with a Todo task list), and `list_plans` retrieves all plans for a group.
#'
#' `get_team` retrieves a team. The method for the Graph client class requires the team ID. The method for the `az_user` class requires either the team name or ID. The method for the `az_group` class retrieves the team associated with the group, if it exists.
#'
#' `get_chat` retrieves a one-on-one, group or meeting chat, by ID. `list_chats` retrieves all chats that the user is part of.
#'
#' Note that Teams, SharePoint and OneDrive for Business require a Microsoft 365 Business license, and are available for organisational tenants only. Similarly, only Microsoft 365 groups can have associated sites/teams/plans/drives, not any other kind of group.
#'
#' @section List methods:
#' All `list_*` methods have `filter` and `n` arguments to limit the number of results. The former should be an [OData expression](https://learn.microsoft.com/en-us/graph/query-parameters#filter-parameter) as a string to filter the result set on. The latter should be a number setting the maximum number of (filtered) results to return. The default values are `filter=NULL` and `n=Inf`. If `n=NULL`, the `ms_graph_pager` iterator object is returned instead to allow manual iteration over the results.
#'
#' Support in the underlying Graph API for OData queries is patchy. Not all endpoints that return lists of objects support filtering, and if they do, they may not allow all of the defined operators. If your filtering expression results in an error, you can carry out the operation without filtering and then filter the results on the client side.
#' @section Value:
#' For `get_sharepoint_site`, an object of class `ms_site`.
#'
#' For `get_drive`, an object of class `ms_drive`. For `list_drives`, a list of `ms_drive` objects.
#'
#' For `get_plan`, an object of class `ms_plan`. For `list_plans`, a list of `ms_plan` objects.
#'
#' For `get_team`, an object of class `ms_team`. For `list_teams`, a list of `ms_team` objects.
#'
#' For `get_chat`, an object of class `ms_chat`. For `list_chats`, a list of `ms_chat` objects.
#' @seealso
#' [`ms_site`], [`ms_drive`], [`ms_plan`], [`ms_team`], [`ms_chat`], [`az_user`], [`az_group`]
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
#' gr2$get_sharepoint_site("My site")
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

add_object_methods <- function()
{
    AzureGraph::ms_object$set("private", "make_basic_list", overwrite=TRUE,
    function(op, filter, n, ...)
    {
        opts <- list(`$filter`=filter)
        hdrs <- if(!is.null(filter)) httr::add_headers(consistencyLevel="eventual")
        pager <- self$get_list_pager(self$do_operation(op, options=opts, hdrs), ...)
        extract_list_values(pager, n)
    })
}

add_graph_methods <- function()
{
    ms_graph$set("public", "get_sharepoint_site", overwrite=TRUE,
    function(site_url=NULL, site_id=NULL)
    {
        assert_one_arg <- get("assert_one_arg", getNamespace("Microsoft365R"))
        assert_one_arg(site_url, site_id, msg="Supply exactly one of site URL or ID")
        op <- if(!is.null(site_url))
        {
            site_url <- httr::parse_url(site_url)
            file.path("sites", paste0(site_url$hostname, ":"), site_url$path)
        }
        else file.path("sites", site_id)
        Microsoft365R::ms_site$new(self$token, self$tenant,
                                   self$call_graph_endpoint(op))
    })

    ms_graph$set("public", "get_drive", overwrite=TRUE,
    function(drive_id)
    {
        op <- file.path("drives", drive_id)
        Microsoft365R::ms_drive$new(self$token, self$tenant,
                                    self$call_graph_endpoint(op))
    })

    ms_graph$set("public", "get_team", overwrite=TRUE,
    function(team_id)
    {
        op <- file.path("teams", team_id)
        Microsoft365R::ms_team$new(self$token, self$tenant,
                                   self$call_graph_endpoint(op))
    })
}

add_user_methods <- function()
{
    az_user$set("public", "list_drives", overwrite=TRUE,
    function(filter=NULL, n=Inf)
    {
        private$make_basic_list("drives", filter, n)
    })

    az_user$set("public", "get_drive", overwrite=TRUE,
    function(drive_id=NULL)
    {
        op <- if(is.null(drive_id))
            "drive"
        else file.path("drives", drive_id)
        Microsoft365R::ms_drive$new(self$token, self$tenant,
                                    self$do_operation(op))
    })

    az_user$set("public", "list_sharepoint_sites", overwrite=TRUE,
    function(filter=NULL, n=Inf)
    {
        lst <- private$make_basic_list("followedSites", filter, n)
        if(!is.null(n))
            lapply(lst, function(site) site$sync_fields())  # result from endpoint is incomplete
        else lst
    })

    az_user$set("public", "list_teams", overwrite=TRUE,
    function(filter=NULL, n=Inf)
    {
        lst <- private$make_basic_list("joinedTeams", filter, n)
        if(!is.null(n))
            lapply(lst, function(team) team$sync_fields())  # result from endpoint only contains ID and displayname
        else lst
    })

    az_user$set("public", "get_outlook", overwrite=TRUE,
    function()
    {
        Microsoft365R::ms_outlook$new(self$token, self$tenant, self$properties)
    })

    az_user$set("public", "list_chats", overwrite=TRUE,
    function(filter=NULL, n=Inf)
    {
        private$make_basic_list("chats", filter, n)
    })

    az_user$set("public", "get_chat", overwrite=TRUE,
    function(chat_id)
    {
        op <- file.path("chats", chat_id)
        Microsoft365R::ms_chat$new(self$token, self$tenant, self$do_operation(op))
    })
}

add_group_methods <- function()
{
    az_group$set("public", "get_sharepoint_site", overwrite=TRUE,
    function()
    {
        res <- self$do_operation("sites/root")
        Microsoft365R::ms_site$new(self$token, self$tenant, res)
    })

    az_group$set("public", "list_drives", overwrite=TRUE,
    function(filter=NULL, n=Inf)
    {
        private$make_basic_list("drives", filter, n)
    })

    az_group$set("public", "get_drive", overwrite=TRUE,
    function(drive_name=NULL, drive_id=NULL)
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
        Microsoft365R::ms_drive$new(self$token, self$tenant,
                                    self$do_operation(op))
    })

    az_group$set("public", "get_team", overwrite=TRUE,
    function()
    {
        op <- file.path("teams", self$properties$id)
        Microsoft365R::ms_team$new(self$token, self$tenant,
                                   call_graph_endpoint(self$token, op))
    })

    az_group$set("public", "list_plans", overwrite=TRUE,
    function(filter=NULL, n=Inf)
    {
        private$make_basic_list("planner/plans", filter, n)
    })

    az_group$set("public", "get_plan", overwrite=TRUE,
    function(plan_title=NULL, plan_id=NULL)
    {
        assert_one_arg(plan_title, plan_id, msg="Supply exactly one of plan title or ID")
        if(!is.null(plan_id))
        {
            res <- call_graph_endpoint(self$token, file.path("planner/plans", plan_id))
            Microsoft365R::ms_plan$new(self$token, self$tenant, res)
        }
        else
        {
            plans <- self$list_plans()
            wch <- which(sapply(plans, function(pl) pl$properties$title == plan_title))
            if(length(wch) != 1)
                stop("Invalid plan title", call.=FALSE)
            plans[[wch]]
        }
    })
}
