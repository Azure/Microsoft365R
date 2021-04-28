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
#' - `delete(confirm=TRUE, by_item=FALSE)`: Delete this item. By default, ask for confirmation first. For personal OneDrive, deleting a folder will also automatically delete its contents; for business OneDrive or SharePoint document libraries, you may need to set `by_item=TRUE` to delete the contents first depending on your organisation's policies. Note that this can be slow for large folders.
#' - `update(...)`: Update the item's properties (metadata) in Microsoft Graph.
#' - `do_operation(...)`: Carry out an arbitrary operation on the item.
#' - `sync_fields()`: Synchronise the R object with the item metadata in Microsoft Graph.
#' - `open()`: Open the item in your browser.
#' - `list_items(...), list_files(...)`: List the files and folders under the specified path.
#' - `download(dest, overwrite)`: Download the file. Only applicable for a file item.
#' - `create_share_link(type, expiry, password, scope)`: Create a shareable link to the file or folder.
#' - `upload(src, dest, blocksize)`: Upload a file. Only applicable for a folder item.
#' - `create_folder(path)`: Create a folder. Only applicable for a folder item.
#' - `get_item(path)`: Get a child item (file or folder) under this folder.
#' - `is_folder()`: Information function, returns TRUE if this item is a folder.
#'
#' @section Initialization:
#' Creating new objects of this class should be done via the `get_item` method of the [`ms_drive`] class. Calling the `new()` method for this class only constructs the R object; it does not call the Microsoft Graph API to retrieve or create the actual item.
#'
#' @section File and folder operations:
#' This class exposes methods for carrying out common operations on files and folders. Note that for the methods below, any paths to child items are relative to the folder's own path.
#'
#' `open` opens this file or folder in your browser. If the file has an unrecognised type, most browsers will attempt to download it.
#'
#' `list_items(path, info, full_names, pagesize)` lists the items under the specified path. It is the analogue of base R's `dir`/`list.files`. Its arguments are
#' - `path`: The path.
#' - `info`: The information to return: either "partial", "name" or "all". If "partial", a data frame is returned containing the name, size, ID and whether the item is a file or folder. If "name", a vector of file/folder names is returned. If "all", a data frame is returned containing _all_ the properties for each item (this can be large).
#' - `full_names`: Whether to prefix the folder path to the names of the items.
#' - `pagesize`: The number of results to return for each call to the REST endpoint. You can try reducing this argument below the default of 1000 if you are experiencing timeouts.
#'
#' `list_files` is a synonym for `list_items`.
#'
#' `download` downloads the file item to the local machine. It is an error to try to download a folder item.
#'
#' `upload` uploads a file from the local machine into the folder item, and returns another `ms_drive_item` object representing the uploaded file. The uploading is done in blocks of 32MB by default; you can change this by setting the `blocksize` argument. For technical reasons, the block size [must be a multiple of 320KB](https://docs.microsoft.com/en-us/graph/api/driveitem-createuploadsession?view=graph-rest-1.0#upload-bytes-to-the-upload-session). This returns an `ms_drive_item` object, invisibly.
#'
#' It is an error to try to upload to a file item, or to upload a source directory.
#'
#' `get_item` retrieves the file or folder with the given path, as another object of class `ms_drive_item`.
#'
#' `create_folder` creates a folder with the specified path. Trying to create an already existing folder is an error. This returns an `ms_drive_item` object, invisibly.
#'
#' `create_share_link(path, type, expiry, password, scope)` returns a shareable link to the item. Its arguments are
#' - `path`: The path.
#' - `type`: Either "view" for a read-only link, "edit" for a read-write link, or "embed" for a link that can be embedded in a web page. The last one is only available for personal OneDrive.
#' - `expiry`: How long the link is valid for. The default is 7 days; you can set an alternative like "15 minutes", "24 hours", "2 weeks", "3 months", etc. To leave out the expiry date, set this to NULL.
#' - `password`: An optional password to protect the link.
#' - `scope`: Optionally the scope of the link, either "anonymous" or "organization". The latter allows only users in your AAD tenant to access the link, and is only available for OneDrive for Business or SharePoint.
#'
#' This method returns a URL to access the item, for `type="view"` or "`type=edit"`. For `type="embed"`, it returns a list with components `webUrl` containing the URL, and `webHtml` containing a HTML fragment to embed the link in an IFRAME. The default is a viewable link, expiring in 7 days.
#'
#' @seealso
#' [`ms_graph`], [`ms_site`], [`ms_drive`]
#'
#' [Microsoft Graph overview](https://docs.microsoft.com/en-us/graph/overview),
#' [OneDrive API reference](https://docs.microsoft.com/en-us/graph/api/resources/onedrive?view=graph-rest-1.0)
#'
#' @examples
#' \dontrun{
#'
#' # personal OneDrive
#' mydrv <- get_personal_onedrive()
#'
#' docs <- mydrv$get_item("Documents")
#' docs$list_files()
#' docs$list_items()
#'
#' # this is the file 'Documents/myfile.docx'
#' myfile <- docs$get_item("myfile.docx")
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

    delete=function(confirm=TRUE, by_item=FALSE)
    {
        if(!by_item || !self$is_folder())
            return(super$delete(confirm=confirm))

        if (confirm && interactive())
        {
            msg <- sprintf("Do you really want to delete the %s '%s'?", self$type, self$properties$name)
            if (!get_confirmation(msg, FALSE))
                return(invisible(NULL))
        }

        children <- self$list_items()
        dirs <- children$isdir
        for(d in children$name[dirs])
            self$get_item(d)$delete(confirm=FALSE, by_item=TRUE)

        deletes <- lapply(children$name[!dirs], function(f)
        {
            path <- private$make_absolute_path(f)
            graph_request$new(path, http_verb="DELETE")
        })
        # do in batches of 20
        i <- length(deletes)
        while(i > 0)
        {
            batch <- seq(from=max(1, i - 19), to=i)
            call_batch_endpoint(self$token, deletes[batch])
            i <- max(1, i - 19) - 1
        }

        super$delete(confirm=FALSE)
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

    list_items=function(path="", info=c("partial", "name", "all"), full_names=FALSE, pagesize=1000, filter=NULL, n=Inf)
    {
        private$assert_is_folder()
        if(path == "/")
            path <- ""
        info <- match.arg(info)
        opts <- switch(info,
            partial=list(`$select`="name,size,folder,id", `$top`=pagesize),
            name=list(`$select`="name", `$top`=pagesize),
            list(`$top`=pagesize)
        )
        if(!is.null(filter))
            opts$`filter` <- filter

        op <- sub("::", "", paste0(private$make_absolute_path(path), ":/children"))
        children <- call_graph_endpoint(self$token, op, options=opts, simplify=TRUE)

        # get file list as a data frame, or return the iterator immediately if n is NULL
        df <- extract_list_values(self$get_list_pager(children), n)
        if(is.null(n))
            return(df)

        if(is_empty(df))
            df <- data.frame(name=character(), size=numeric(), isdir=logical(), id=character())
        else if(info != "name")
        {
            df$isdir <- if(!is.null(df$folder))
                !is.na(df$folder$childCount)
            else rep(FALSE, nrow(df))
        }

        if(full_names)
            df$name <- file.path(sub("^/", "", path), df$name)
        switch(info,
            partial=df[c("name", "size", "isdir", "id")],
            name=df$name,
            all=
            {
                firstcols <- c("name", "size", "isdir", "id")
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

        # see https://stackoverflow.com/a/66686842/474349
        op <- private$make_absolute_path(path)
        body <- list(
            folder=named_list(),
            `@microsoft.graph.conflictBehavior`="fail"
        )
        res <- call_graph_endpoint(self$token, op, body=body, http_verb="PATCH")
        invisible(ms_drive_item$new(self$token, self$tenant, res))
    },

    upload=function(src, dest=basename(src), blocksize=32768000)
    {
        private$assert_is_folder()
        con <- file(src, open="rb")
        on.exit(close(con))

        op <- paste0(private$make_absolute_path(dest), ":/createUploadSession")
        # print(op)
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
                    next_blockstart, next_blockstart + thisblock - 1, size)
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
        if(dest == ".")
            dest <- ""
        parent <- self$properties$parentReference
        name <- self$properties$name
        op <- if(name == "root")
            file.path("drives", parent$driveId, "root:")
        else
        {
            # have to infer the parent path if we got this item as a Teams channel folder
            # in this case, assume the parent is the root folder
            if(is.null(parent$path))
                parent$path <- sprintf("/drives/%s/root:", parent$driveId)
            file.path(parent$path, name)
        }
        utils::URLencode(enc2utf8(sub("/$", "", file.path(op, dest))))
    },

    assert_is_folder=function()
    {
        if(!self$is_folder())
            stop("This method is only applicable for a folder item", call.=FALSE)
    },

    assert_is_file=function()
    {
        if(self$is_folder())
            stop("This method is only applicable for a file item", call.=FALSE)
    }
))


# alias for convenience
ms_drive_item$set("public", "list_files", overwrite=TRUE, ms_drive_item$public_methods$list_items)
