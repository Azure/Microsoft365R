#' @import AzureGraph
NULL

utils::globalVariables(c("self", "private"))

.onLoad <- function(libname, pkgname)
{
    register_graph_class("site", ms_site,
        function(props) grepl("sharepoint", props$id, fixed=TRUE))

    register_graph_class("drive", ms_drive,
        function(props) !is_empty(props$driveType) && is_empty(props$parentReference))

    register_graph_class("driveItem", ms_drive_item,
        function(props) !is_empty(props$parentReference$driveId))

    register_graph_class("list", ms_sharepoint_list,
        function(props) !is_empty(props$list))

    add_methods()
}

# authentication app ID
.microsoft365r_app_id <- "d44a05d5-c6a5-4bbb-82d2-443123722380"

# helper function
error_message <- get("error_message", getNamespace("AzureGraph"))

# dummy mention to keep CRAN happy
# we need to ensure that vctrs is loaded so that AzureGraph will use vec_rbind
# to combine paged results into a single data frame: individual pages can have
# different structures, which will break base::rbind
vctrs::vec_rbind
