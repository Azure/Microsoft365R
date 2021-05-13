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
#'
#' @section Initialization:
#' Creating new objects of this class should be done via the `list_buckets` method of the [`ms_plan`] class.
#' Calling the `new()` method for this class only constructs the R object; it does not call the Microsoft Graph API to retrieve or create the actual plan bucket.
#'
#' @section Plan bucket operations:
#' This class exposes methods for carrying out common operations on a plan bucket.
#'
#' [Microsoft Graph overview](https://docs.microsoft.com/en-us/graph/overview),
#' [OneDrive API reference](https://docs.microsoft.com/en-us/graph/api/resources/planner?view=graph-rest-1.0)
#'
#' @examples
#' \dontrun{
#' }
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

    print=function(...)
    {
        name <- paste0("<Bucket ", self$properties$name, ">\n")
        cat(name)
        cat("---\n")
        cat(format_public_methods(self))
        invisible(self)
    }
))
