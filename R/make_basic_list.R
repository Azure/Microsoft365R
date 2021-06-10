#' Create a list of Graph objects
#' @param object An R6 object inheriting from `AzureGraph::ms_object`.
#' @param op A string, the REST operation.
#' @param filter,n Filtering arguments for the list.
#' @param ... Arguments passed to lower-level functions, and ultimately to `AzureGraph::call_graph_endpoint`.
#' @details
#' This function is a basic utility called by various Microsoft365R class methods. It is exported to work around issues in how R6 handles classes that extend across package boundaries. It should not be called by the user.
#' @return
#' If `n` is NULL, an iterator object, of class `AzureGraph::ms_graph_pager`. Otherwise, a list of individual Graph objects.
#' @seealso
#' [`AzureGraph::call_graph_endpoint`], [`AzureGraph::ms_graph_pager`]
#' @export
make_basic_list <- function(object, op, filter, n, ...)
{
    opts <- list(`$filter`=filter)
    hdrs <- if(!is.null(filter)) httr::add_headers(consistencyLevel="eventual")
    pager <- object$get_list_pager(object$do_operation(op, options=opts, hdrs), ...)
    extract_list_values(pager, n)
}

