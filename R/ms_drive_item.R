#' File or folder in a drive
#'
#' Class representing an item (file or folder) in a OneDrive or SharePoint document library.
#'
#' @docType class
#' @section Fields:
#' - `token`: The token used to authenticate with the Graph host.
#' - `tenant`: The Azure Active Directory tenant for the parent drive.
#' - `type`: always "drive item" for a drive item object.
#' - `properties`: The item properties (metadata).
#' @section Methods:
#' - `new(...)`: Initialize a new object. Do not call this directly; see 'Initialization' below.
#' - `delete(confirm=TRUE)`: Delete this item. By default, ask for confirmation first.
#' - `update(...)`: Update the item's properties (metadata) in Microsoft Graph.
#' - `do_operation(...)`: Carry out an arbitrary operation on the item.
#' - `sync_fields()`: Synchronise the R object with the item metadata in Microsoft Graph.
#' - `open()`: Open the item in your browser.
#' - `download(dest, overwrite)`: Download the file. Not applicable for a folder.
#' - `create_share_link(type, expiry, password, scope)`: Create a shareable link to the file or folder. See 'Sharing' below.
#'
#' @section Initialization:
#' Creating new objects of this class should be done via the `get_item` method of the [ms_drive] class. Calling the `new()` method for this class only constructs the R object; it does not call the Microsoft Graph API to retrieve or create the actual item.
#'
#' @section Sharing:
#' `create_share_link(type, expiry, password, scope)` returns a shareable link to the item. Its arguments are
#' - `type`: Either "view" for a read-only link, "edit" for a read-write link, or "embed" for a link that can be embedded in a web page. The last one is only available for personal OneDrive.
#' - `expiry`: How long the link is valid for. The default is 7 days; you can set an alternative like "15 minutes", "24 hours", "2 weeks", "3 months", etc. To leave out the expiry date, set this to NULL.
#' - `password`: An optional password to protect the link.
#' - `scope`: Optionally the scope of the link, either "anonymous" or "organization". The latter allows only users in your AAD tenant to access the link, and is only available for OneDrive for Business or SharePoint.
#'
#' This function returns a URL to access the item, for `type="view"` or "`type=edit"`. For `type="embed"`, it returns a list with components `webUrl` containing the URL, and `webHtml` containing a HTML fragment to embed the link in an IFRAME.
#' @seealso
#' [ms_graph], [ms_site], [ms_drive]
#'
#' [Microsoft Graph overview](https://docs.microsoft.com/en-us/graph/overview),
#' [OneDrive API reference](https://docs.microsoft.com/en-us/graph/api/resources/onedrive?view=graph-rest-1.0)
#'
#' @examples
#' \dontrun{
#'
#' # personal OneDrive
#' gr2 <- get_graph_login("consumers")
#' me <- gr2$get_user()
#' mydrv <- me$get_drive()
#'
#' myfile <- drv$get_item("myfile.docx")
#' myfile$properties
#'
#' # rename a file
#' myfile$update(name="newname.docx")
#'
#' # open the file in the browser
#' myfile$open()
#'
#' # download the file to the working directory
#' myfile$download()
#'
#' # shareable links
#' myfile$create_share_link()
#' myfile$create_share_link(type="edit", expiry="24 hours")
#' myfile$create_share_link(password="Use-strong-passwords!")
#'
#' # delete the file (will ask for confirmation first)
#' myfile$delete()
#'
#' }
#' @format An R6 object of class `ms_drive_item`, inheriting from `ms_object`.
#' @export
ms_drive_item <- R6::R6Class("ms_drive_item", inherit=ms_object,

public=list(

    initialize=function(token, tenant=NULL, properties=NULL)
    {
        self$type <- "drive item"
        private$api_type <- file.path("drives", properties$parentReference$driveId, "items")
        super$initialize(token, tenant, properties)
    },

    is_folder=function()
    {
        !is.null(self$properties$folder)
    },

    open=function()
    {
        httr::BROWSE(self$properties$webUrl)
    },

    create_share_link=function(type=c("view", "edit", "embed"), expiry="7 days", password=NULL, scope=NULL)
    {
        type <- match.arg(type)
        body <- list(type=type)
        if(!is.null(expiry))
        {
            expdate <- seq(Sys.time(), by=expiry, len=2)[2]
            expirationDateTime <- strftime(expdate, "%Y-%m-%dT%H:%M:%SZ", tz="GMT")
        }
        if(!is.null(password))
            body$password <- password
        if(!is.null(scope))
            body$scope <- scope
        res <- self$do_operation("createLink", body=body, http_verb="POST")
        if(type == "embed")
        {
            res$link$type <- NULL
            res$link
        }
        else res$link$webUrl
    },

    list_items=function(path="", info=c("partial", "name", "all"), full_names=FALSE, pagesize=1000)
    {
        private$assert_is_folder()
        info <- match.arg(info)
        opts <- switch(info,
            partial=list(`$select`="name,size,folder", `$top`=pagesize),
            name=list(`$select`="name", `$top`=pagesize),
            list(`$top`=pagesize)
        )

        op <- sub("::", "", paste0(private$make_absolute_path(path), ":/children"))
        children <- call_graph_endpoint(self$token, op, options=opts, simplify=TRUE)

        # get file list as a data frame
        df <- private$get_paged_list(children, simplify=TRUE)

        if(is_empty(df))
            df <- data.frame(name=character(), size=numeric(), isdir=logical())
        else if(info != "name")
        {
            df$isdir <- if(!is.null(df$folder))
                !is.na(df$folder$childCount)
            else rep(FALSE, nrow(df))
        }

        if(full_names)
            df$name <- file.path(sub("^/", "", path), df$name)
        switch(info,
            partial=df[c("name", "size", "isdir")],
            name=df$name,
            all=
            {
                firstcols <- c("name", "size", "isdir")
                df[c(firstcols, setdiff(names(df), firstcols))]
            }
        )
    },

    get_item=function(path)
    {
        private$assert_is_folder()
        op <- private$make_absolute_path(path)
        ms_drive_item$new(self$token, self$tenant, call_graph_endpoint(self$token, op))
    },

    create_folder=function(path)
    {
        private$assert_is_folder()
        body <- list(
            name=enc2utf8(path),
            folder=named_list(),
            `@microsoft.graph.conflictBehavior`="fail"
        )
        res <- self$do_operation("children", body=body, http_verb="POST")
        invisible(ms_drive_item$new(self$token, self$tenant, res))
    },

    upload=function(src, dest=basename(src), blocksize=32768000)
    {
        private$assert_is_folder()
        con <- file(src, open="rb")
        on.exit(close(con))

        op <- paste0(private$make_absolute_path(dest), ":/createUploadSession")
        upload_dest <- call_graph_endpoint(self$token, op, http_verb="POST")$uploadUrl

        size <- file.size(src)
        next_blockstart <- 0
        next_blockend <- size - 1
        repeat
        {
            next_blocksize <- min(next_blockend - next_blockstart + 1, blocksize)
            seek(con, next_blockstart)
            body <- readBin(con, "raw", next_blocksize)
            thisblock <- length(body)
            if(thisblock == 0)
                break

            headers <- httr::add_headers(
                `Content-Length`=thisblock,
                `Content-Range`=sprintf("bytes %.0f-%.0f/%.0f",
                    next_blockstart, next_blockstart + next_blocksize - 1, size)
            )
            res <- httr::PUT(upload_dest, headers, body=body)
            httr::stop_for_status(res)

            next_block <- parse_upload_range(httr::content(res), blocksize)
            if(is.null(next_block))
                break
            next_blockstart <- next_block[1]
            next_blockend <- next_block[2]
        }
        invisible(ms_drive_item$new(self$token, self$tenant, httr::content(res)))
    },

    download=function(dest=self$properties$name, overwrite=FALSE)
    {
        private$assert_is_file()
        filepath <- file.path(self$parentReference$path, self$properties$name)
        res <- self$do_operation("content", config=httr::write_disk(dest, overwrite=overwrite),
                                 http_status_handler="pass")
        if(httr::status_code(res) >= 300)
        {
            on.exit(file.remove(dest))
            httr::stop_for_status(res, paste0("complete operation. Message:\n",
                sub("\\.$", "", error_message(httr::content(res)))))
        }
        invisible(NULL)
    },

    print=function(...)
    {
        file_or_dir <- if(self$is_folder()) "file folder" else "file"
        cat("<Drive item '", self$properties$name, "'>\n", sep="")
        cat("  directory id:", self$properties$id, "\n")
        cat("  web link:", self$properties$webUrl, "\n")
        cat("  type:", file_or_dir, "\n")
        cat("---\n")
        cat(format_public_methods(self))
        invisible(self)
    }
),

private=list(

    make_absolute_path=function(dest)
    {
        parent <- self$properties$parentReference
        name <- self$properties$name
        op <- if(name == "root")
            file.path("drives", parent$driveId, "root:")
        else file.path(parent$path, name)
        file.path(op, enc2utf8(dest))
    },

    assert_is_folder=function()
    {
        if(!self$is_folder())
            stop("This method must be called on a folder", call.=FALSE)
    },

    assert_is_file=function()
    {
        if(self$is_folder())
            stop("This method must be called on a file", call.=FALSE)
    }
))


# alias for convenience
ms_drive_item$set("public", "list_files", overwrite=TRUE, ms_drive_item$public_methods$list_items)
