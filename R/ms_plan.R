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
#' - `list_tasks(filter=NULL, n=Inf)`: List the tasks under the specified plan.
#' - `get_task(task_title, task_id)`: Get a task, either by title or ID.
#' - `list_buckets(filter=NULL, n=Inf)`: List the buckets under the specified plan.
#' - `get_bucket(bucket_name, bucket_id)`: Get a bucket, either by name or ID.
#' - `get_details()`: Get the plan details.
#'
#' @section Initialization:
#' Creating new objects of this class should be done via the `list_plans` methods of the [`az_group`] class.
#' Calling the `new()` method for this class only constructs the R object; it does not call the Microsoft Graph API to retrieve or create the actual plan.
#'
#' @section Planner operations:
#' This class exposes methods for carrying out common operations on a plan. Currently only read operations are supported.
#'
#' Call `list_tasks()` to list the tasks under the plan, and `get_task()` to retrieve a specific task. Similarly, call `list_buckets()` to list the buckets, and `get_bucket()` to retrieve a specific bucket.
#'
#' Call `get_details()` to get a list containing the details for the plan.
#'
#' @section List methods:
#' All `list_*` methods have `filter` and `n` arguments to limit the number of results. The former should be an [OData expression](https://learn.microsoft.com/en-us/graph/query-parameters#filter-parameter) as a string to filter the result set on. The latter should be a number setting the maximum number of (filtered) results to return. The default values are `filter=NULL` and `n=Inf`. If `n=NULL`, the `ms_graph_pager` iterator object is returned instead to allow manual iteration over the results.
#'
#' Support in the underlying Graph API for OData queries is patchy. Not all endpoints that return lists of objects support filtering, and if they do, they may not allow all of the defined operators. If your filtering expression results in an error, you can carry out the operation without filtering and then filter the results on the client side.
#' @seealso
#' [`ms_plan_task`], [`ms_plan_bucket`]
#'
#' [Microsoft Graph overview](https://learn.microsoft.com/en-us/graph/overview),
#' [Plans overview](https://learn.microsoft.com/en-us/graph/api/resources/planner-overview?view=graph-rest-beta)
#'
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

    list_tasks=function(filter=NULL, n=Inf)
    {
        private$make_basic_list("tasks", filter, n)
    },

    get_task=function(task_title=NULL, task_id=NULL)
    {
        assert_one_arg(task_title, task_id, msg="Supply exactly one of task title or ID")
        if(!is.null(task_id))
        {
            res <- call_graph_endpoint(self$token, file.path("planner/tasks", task_id))
            ms_plan_task$new(self$token, self$tenant, res)
        }
        else
        {
            tasks <- self$list_tasks(filter=sprintf("title eq '%s'", task_title))
            if(length(tasks) != 1)
                stop("Invalid task title", call.=FALSE)
            tasks[[1]]
        }
    },

    list_buckets=function(filter=NULL, n=Inf)
    {
        private$make_basic_list("buckets", filter, n)
    },

    get_bucket=function(bucket_name=NULL, bucket_id=NULL)
    {
        assert_one_arg(bucket_name, bucket_id, msg="Supply exactly one of bucket name or ID")
        if(!is.null(bucket_id))
        {
            res <- call_graph_endpoint(self$token, file.path("planner/buckets", bucket_id))
            ms_plan_bucket$new(self$token, self$tenant, res)
        }
        else
        {
            buckets <- self$list_buckets(filter=sprintf("name eq '%s'", bucket_name))
            if(length(buckets) != 1)
                stop("Invalid bucket name", call.=FALSE)
            buckets[[1]]
        }
    },

    get_details=function()
    {
        self$do_operation("details")
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
