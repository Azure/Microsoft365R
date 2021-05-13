#' Microsoft Planner Plan
#'
#' Class representing one plan withing a Microsoft Planner.
#'
#' The plans belong to a group.
#'
#' @docType class
#' @section Fields:
#' - `token`: The token used to authenticate with the Graph host.
#' - `tenant`: The Azure Active Directory tenant for this drive.
#' - `type`: always "plan" for plan object.
#' - `properties`: The plan properties.
#' @section Methods:
#' - `new(...)`: Initialize a new plan object. Do not call this directly; see 'Initialization' below.
#' - `update(...)`: Update the plan metadata in Microsoft Graph.
#' - `do_operation(...)`: Carry out an arbitrary operation on the plan
#' - `sync_fields()`: Synchronise the R object with the plan metadata in Microsoft Graph.
#' - `list_tasks(...)`: List the tasks under the specified plan.
#'
#' @section Initialization:
#' Creating new objects of this class should be done via the `list_plans` methods of the [`az_group`] class.
#' Calling the `new()` method for this class only constructs the R object; it does not call the Microsoft Graph API to retrieve or create the actual plan.
#'
#' @section Planner operations:
#' This class exposes methods for carrying out common operations on a plan.
#'
#' `list_tasks()` lists the tasks under the plan.
#'
#' @seealso
#' [`ms_plan_task`]
#'
#' [Microsoft Graph overview](https://docs.microsoft.com/en-us/graph/overview),
#' [OneDrive API reference](https://docs.microsoft.com/en-us/graph/api/resources/planner?view=graph-rest-1.0)
#'
#' @examples
#' \dontrun{
#' }
#' @format An R6 object of class `ms_plan`, inheriting from `ms_object`.
#' @export
ms_plan <- R6::R6Class("ms_plan", inherit=ms_object,

public=list(

    initialize=function(token, tenant=NULL, properties=NULL)
    {
        self$type <- "plan"
        private$api_type <- "planner/plans"
        super$initialize(token, tenant, properties)
    },

    list_tasks=function()
    {
        res <- private$get_paged_list(self$do_operation("tasks"))
        private$init_list_objects(res, "plan_task")
    },

    list_buckets=function()
    {
        res <- private$get_paged_list(self$do_operation("buckets"))
        private$init_list_objects(res, "plan_bucket")
    },

    print=function(...)
    {
        name <- paste0("<Plan for ", self$properties$title, ">\n")
        cat(name)
        cat("---\n")
        cat(format_public_methods(self))
        invisible(self)
    }
))
