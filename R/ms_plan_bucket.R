#' Microsoft Planner Plan Bucket
#'
#' Class representing a bucket within a plan of a Microsoft Planner.
#'
#' @docType class
#' @section Fields:
#' - `token`: The token used to authenticate with the Graph host.
#' - `tenant`: The Azure Active Directory tenant for this bucket
#' - `type`: always "plan_bucket" for plan bucket object.
#' - `properties`: The plan bucket properties.
#' @section Methods:
#' - `new(...)`: Initialize a new plan bucket object. Do not call this directly; see 'Initialization' below.
#' - `update(...)`: Update the plan bucket metadata in Microsoft Graph.
#' - `do_operation(...)`: Carry out an arbitrary operation on the plan bucket
#' - `sync_fields()`: Synchronise the R object with the plan bucket metadata in Microsoft Graph.
#' - `list_tasks(filter=NULL, n=Inf)`: List the tasks for this bucket.
#'
#' @section Initialization:
#' Creating new objects of this class should be done via the `list_buckets` method of the [`ms_plan`] class.
#' Calling the `new()` method for this class only constructs the R object; it does not call the Microsoft Graph API to retrieve or create the actual plan bucket.
#'
#' @section List methods:
#' All `list_*` methods have `filter` and `n` arguments to limit the number of results. The former should be an [OData expression](https://learn.microsoft.com/en-us/graph/query-parameters#filter-parameter) as a string to filter the result set on. The latter should be a number setting the maximum number of (filtered) results to return. The default values are `filter=NULL` and `n=Inf`. If `n=NULL`, the `ms_graph_pager` iterator object is returned instead to allow manual iteration over the results.
#'
#' Support in the underlying Graph API for OData queries is patchy. Not all endpoints that return lists of objects support filtering, and if they do, they may not allow all of the defined operators. If your filtering expression results in an error, you can carry out the operation without filtering and then filter the results on the client side.
#' @seealso
#' [`ms_plan`], [`ms_plan_task`]
#'
#' [Microsoft Graph overview](https://learn.microsoft.com/en-us/graph/overview),
#' [Plans overview](https://learn.microsoft.com/en-us/graph/api/resources/planner-overview?view=graph-rest-beta)
#'
#' @format An R6 object of class `ms_plan_bucket`, inheriting from `ms_object`.
#' @export
ms_plan_bucket <- R6::R6Class("ms_plan_bucket", inherit=ms_object,

public=list(

    initialize=function(token, tenant=NULL, properties=NULL)
    {
        self$type <- "plan_bucket"
        private$api_type <- "planner/buckets"
        super$initialize(token, tenant, properties)
    },

    list_tasks=function(filter=NULL, n=Inf)
    {
        private$make_basic_list("tasks", filter, n)
    },

    print=function(...)
    {
        name <- paste0("<Bucket ", self$properties$name, ">\n")
        cat(name)
        cat("---\n")
        cat(format_public_methods(self))
        invisible(self)
    }
))
