ms_channel <- R6::R6Class("ms_channel", inherit=ms_object,

public=list(

    team_id=NULL,

    initialize=function(token, tenant=NULL, properties=NULL, team_id)
    {
        self$type <- "channel"
        self$team_id <- team_id
        gid <- parse_channel_weburl(properties$webUrl)
        private$api_type <- file.path("teams", self$team_id, "channels")
        super$initialize(token, tenant, properties)
    },

    send_message=function(body, ...)
    {
        call_body <- list(body=body, ...)
        res <- self$do_operation("messages", body=call_body, http_verb="POST")
        chat_message$new(self$token, self$tenant, res,
                         parent=list(team_id=self$team_id, channel_id=self$properties$id))
    },

    list_messages=function(n=50)
    {
        lst <- private$get_paged_list(self$do_operation("messages"), n=n)
        private$init_list_objects(lst, "chatMessage", parent=list(team_id=self$team_id, channel_id=self$properties$id))
    },

    get_message=function(message_id)
    {
        op <- file.path("messages", message_id)
        ms_chat_message$new(self$token, self$tenant, self$do_operation(op),
                            parent=list(team_id=self$team_id, channel_id=self$properties$id))
    },

    delete_message=function(message_id, confirm=TRUE)
    {
        self$get_message(message_id)$delete(confirm=confirm)
    },

    list_files=function(path="", ...)
    {
        path <- sub("/$", "", file.path(self$properties$displayName, path))
        private$get_drive()$list_files(path, ...)
    },

    download_file=function(src, dest=basename(src), ...)
    {
        src <- file.path(self$properties$displayName, src)
        private$get_drive()$download_file(src, dest, ...)
    },

    upload_file=function(src, dest, ...)
    {
        dest <- file.path(self$properties$displayName, dest)
        private$get_drive()$upload_file(src, dest, ...)
    },

    print=function(...)
    {
        cat("<Teams channel '", self$properties$displayName, "'>\n", sep="")
        cat("  directory id:", self$properties$id, "\n")
        cat("  web link:", self$properties$webUrl, "\n")
        cat("---\n")
        cat(format_public_methods(self))
        invisible(self)
    }
),

private=list(

    get_drive=function(drive_id=NULL)
    {
        op <- if(is.null(drive_id))
            "drive"
        else file.path("drives", drive_id)
        ms_drive$new(self$token, self$tenant, private$do_group_operation(op))
    },

    get_group=function()
    {
        az_group$new(self$token, self$tenant, private$do_group_operation())
    },

    do_group_operation=function(op="", ...)
    {
        gid <- parse_channel_weburl(self$properties$webUrl)
        op <- sub("/$", "", file.path("groups", gid, op))
        call_graph_endpoint(self$token, op, ...)
    }

    # get_paged_list=function(lst, next_link_name="@odata.nextLink", value_name="value", simplify=FALSE, n=Inf)
    # {
    #     res <- lst[[value_name]]
    #     if(n <= 0) n <- Inf
    #     while(!is_empty(lst[[next_link_name]]) && length(res) < n)
    #     {
    #         lst <- call_graph_url(self$token, lst[[next_link_name]], simplify=simplify)
    #         res <- if(simplify)
    #             vctrs::vec_rbind(res, lst[[value_name]])
    #         else c(res, lst[[value_name]])
    #     }
    #     if(n < length(res))
    #         res[seq_len(n)]
    #     else res
    # }
))


parse_channel_weburl <- function(x)
{
    if(is.null(x))
        stop("Unable to initialize team channel object: no web URL", call.=FALSE)
    httr::parse_url(x)$query$groupId
}
