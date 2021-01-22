#' @import AzureGraph
NULL

utils::globalVariables(c("self", "private"))

.onLoad <- function(libname, pkgname)
{
    # set Graph API to beta, for more powerful permissions
    options(azure_graph_api_version="beta")

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

# default app ID
.microsoft365r_app_id <- "d44a05d5-c6a5-4bbb-82d2-443123722380"

# CLI for Microsoft 365 app ID
.cli_microsoft365_app_id <- "31359c7f-bd7e-475c-86db-fdb8c937548e"

# helper function
error_message <- get("error_message", getNamespace("AzureGraph"))

# dummy mention to keep CRAN happy
# we need to ensure that vctrs is loaded so that AzureGraph will use vec_rbind
# to combine paged results into a single data frame: individual pages can have
# different structures, which will break base::rbind
vctrs::vec_rbind
