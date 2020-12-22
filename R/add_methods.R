add_methods <- function()
{
    ms_graph$set("public", "get_sharepoint_site", overwrite=TRUE,
    function(site_url=NULL, site_id=NULL)
    {
        op <- if(is.null(site_url) && !is.null(site_id))
            file.path("sites", site_id)
        else if(!is.null(site_url) && is.null(site_id))
        {
            site_url <- httr::parse_url(site_url)
            file.path("sites", paste0(site_url$hostname, ":"), site_url$path)
        }
        else stop("Must supply either site ID or URL")

        ms_site$new(self$token, self$tenant, self$call_graph_endpoint(op))
    })

    ms_graph$set("public", "get_drive", overwrite=TRUE,
    function(drive_id)
    {
        op <- file.path("drives", drive_id)
        ms_drive$new(self$token, self$tenant, self$call_graph_endpoint(op))
    })

    az_user$set("public", "list_drives", overwrite=TRUE,
    function()
    {
        res <- private$get_paged_list(self$do_operation("drives"))
        private$init_list_objects(res, "drive")
    })

    az_user$set("public", "get_drive", overwrite=TRUE,
    function(drive_id=NULL)
    {
        op <- if(is.null(drive_id))
            "drive"
        else file.path("drives", drive_id)
        ms_drive$new(self$token, self$tenant, self$do_operation(op))
    })

    az_group$set("public", "get_sharepoint_site", overwrite=TRUE,
    function()
    {
        res <- self$do_operation("sites/root")
        ms_site$new(self$token, self$tenant, res)
    })

    az_group$set("public", "list_drives", overwrite=TRUE,
    function()
    {
        res <- private$get_paged_list(self$do_operation("drives"))
        private$init_list_objects(res, "drive")
    })

    az_group$set("public", "get_drive", overwrite=TRUE, function(drive_id=NULL)
    {
        op <- if(is.null(drive_id))
            "drive"
        else file.path("drives", drive_id)
        ms_drive$new(self$token, self$tenant, self$do_operation(op))
    })
}
