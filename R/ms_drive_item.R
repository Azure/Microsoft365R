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
#' - `copy(dest, dest_folder_item=NULL)`: Copy the item to the given location.
#' - `move(dest, dest_folder_item=NULL)`: Move the item to the given location.
#' - `list_items(...), list_files(...)`: List the files and folders under the specified path.
#' - `download(dest, overwrite, recursive, parallel)`: Download the file or folder. See below.
#' - `create_share_link(type, expiry, password, scope)`: Create a shareable link to the file or folder.
#' - `upload(src, dest, blocksize, , recursive, parallel)`: Upload a file or folder. See below.
#' - `create_folder(path)`: Create a folder. Only applicable for a folder item.
#' - `get_item(path)`: Get a child item (file or folder) under this folder.
#' - `get_parent_folder()`: Get the parent folder for this item, as a drive item object. Returns the root folder for the root. Not supported for remote items.
#' - `get_path()`: Get the absolute path for this item, as a character string. Not supported for remote items.
#' - `is_folder()`: Information function, returns TRUE if this item is a folder.
#' - `load_dataframe(delim=NULL, ...)`: Download a delimited file and return its contents as a data frame. See 'Saving and loading data' below.
#' - `load_rds()`: Download a .rds file and return the saved object.
#' - `load_rdata(envir)`: Load a .RData or .Rda file into the specified environment.
#' - `save_dataframe(df, file, delim=",", ...)` Save a dataframe to a delimited file.
#' - `save_rds(object, file)`: Save an R object to a .rds file.
#' - `save_rdata(..., file)`: Save the specified objects to a .RData file.
#'
#' @section Initialization:
#' Creating new objects of this class should be done via the `get_item` method of the [`ms_drive`] class. Calling the `new()` method for this class only constructs the R object; it does not call the Microsoft Graph API to retrieve or create the actual item.
#'
#' @section File and folder operations:
#' This class exposes methods for carrying out common operations on files and folders. Note that for the methods below, any paths to child items are relative to the folder's own path.
#'
#' `open` opens this file or folder in your browser. If the file has an unrecognised type, most browsers will attempt to download it.
#'
#' `list_items(path, info, full_names, filter, n, pagesize)` lists the items under the specified path. It is the analogue of base R's `dir`/`list.files`. Its arguments are
#' - `path`: The path.
#' - `info`: The information to return: either "partial", "name" or "all". If "partial", a data frame is returned containing the name, size, ID and whether the item is a file or folder. If "name", a vector of file/folder names is returned. If "all", a data frame is returned containing _all_ the properties for each item (this can be large).
#' - `full_names`: Whether to prefix the folder path to the names of the items.
#' - `filter, n`: See 'List methods' below.
#' - `pagesize`: The number of results to return for each call to the REST endpoint. You can try reducing this argument below the default of 1000 if you are experiencing timeouts.
#'
#' `list_files` is a synonym for `list_items`.
#'
#' `download` downloads the item to the local machine. If this is a file, it is downloaded; in this case, the `dest` argument can be the path to the destination file, or NULL to return the downloaded content in a raw vector. If the item is a folder, all its files are downloaded, including subfolders if the `recursive` argument is TRUE.
#'
#' `upload` uploads a file or folder from the local machine into the folder item. The `src` argument can be the path to the source file, a [rawConnection] or a [textConnection] object. If `src` is a folder, all its files are uploaded, including subfolders if the `recursive` argument iS TRUE. An `ms_drive_item` object is returned invisibly.
#'
#' Uploading is done in blocks of 32MB by default; you can change this by setting the `blocksize` argument. For technical reasons, the block size [must be a multiple of 320KB](https://learn.microsoft.com/en-us/graph/api/driveitem-createuploadsession?view=graph-rest-1.0#upload-bytes-to-the-upload-session).
#'
#' Uploading and downloading folders can be done in parallel, which can result in substantial speedup when transferring a large number of small files. This is controlled by the `parallel` argument to `upload` and `download`, which can have the following values:
#' - TRUE: A cluster with 5 workers is created
#' - A number: A cluster with this many workers is created
#' - A cluster object, created via the parallel package
#' - FALSE: The transfer is done serially
#'
#' `get_item` retrieves the file or folder with the given path, as another object of class `ms_drive_item`.
#'
#' - `copy` and `move` can take the destination location as either a full pathname (in the `dest` argument), or a name plus a drive item object (in the `dest_folder_item` argument). If the latter is supplied, any path in `dest` is ignored with a warning. Note that copying is an _asynchronous_ operation, meaning the method returns before the copy is complete.
#'
#' For copying and moving, the destination folder must exist beforehand. When copying/moving a large number of files, it's much more efficient to supply the destination folder in the `dest_folder_item` argument rather than as a path.
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
#' @section Saving and loading data:
#' The following methods are provided to simplify the task of loading and saving datasets and R objects.
#' - `load_dataframe` downloads a delimited file and returns its contents as a data frame. The delimiter can be specified with the `delim` argument; if omitted, this is "," if the file extension is .csv, ";" if the file extension is .csv2, and a tab otherwise. If the readr package is installed, the `readr::read_delim` function is used to parse the file, otherwise `utils::read.delim` is used. You can supply other arguments to the parsing function via the `...` argument.
#' - `save_dataframe` is the inverse of `load_dataframe`: it uploads the given data frame to a folder item. Specify the delimiter with the `delim` argument. The `readr::write_delim` function is used to serialise the data if that package is installed, and `utils::write.table` otherwise.
#' - `load_rds` downloads a .rds file and returns its contents as an R object. It is analogous to the base `readRDS` function but for OneDrive/SharePoint drive items.
#' - `save_rds` uploads a given R object as a .rds file, analogously to `saveRDS`.
#' - `load_rdata` downloads a .RData or .Rda file and loads its contents into the given environment. It is analogous to the base `load` function but for OneDrive/SharePoint drive items.
#' - `save_rdata` uploads the given R objects as a .RData file, analogously to `save`.
#'
#' @section List methods:
#' All `list_*` methods have `filter` and `n` arguments to limit the number of results. The former should be an [OData expression](https://learn.microsoft.com/en-us/graph/query-parameters#filter-parameter) as a string to filter the result set on. The latter should be a number setting the maximum number of (filtered) results to return. The default values are `filter=NULL` and `n=Inf`. If `n=NULL`, the `ms_graph_pager` iterator object is returned instead to allow manual iteration over the results.
#'
#' Support in the underlying Graph API for OData queries is patchy. Not all endpoints that return lists of objects support filtering, and if they do, they may not allow all of the defined operators. If your filtering expression results in an error, you can carry out the operation without filtering and then filter the results on the client side.
#' @seealso
#' [`ms_graph`], [`ms_site`], [`ms_drive`]
#'
#' [Microsoft Graph overview](https://learn.microsoft.com/en-us/graph/overview),
#' [OneDrive API reference](https://learn.microsoft.com/en-us/graph/api/resources/onedrive?view=graph-rest-1.0)
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
#' # copy a file (destination folder must exist)
#' myfile$copy("/Documents/folder2/myfile_copied.docx")
#'
#' # alternate way of copying: supply the destination folder
#' destfolder <- docs$get_item("folder2")
#' myfile$copy("myfile_copied.docx", dest_folder_item=destfolder)
#'
#' # move a file (destination folder must exist)
#' myfile$move("Documents/folder2/myfile_moved.docx")
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
#' # saving and loading data
#' myfolder <- mydrv$get_item("myfolder")
#' myfolder$save_dataframe(iris, "iris.csv")
#' iris2 <- myfolder$get_item("iris.csv")$load_dataframe()
#' identical(iris, iris2)  # TRUE
#'
#' myfolder$save_rds(iris, "iris.rds")
#' iris3 <- myfolder$get_item("iris.rds")$load_rds()
#' identical(iris, iris3)  # TRUE
#'
#' }
#' @format An R6 object of class `ms_drive_item`, inheriting from `ms_object`.
#' @export
ms_drive_item <- R6::R6Class("ms_drive_item", inherit=ms_object,

public=list(

    initialize=function(token, tenant=NULL, properties=NULL, remote=NULL)
    {
        self$type <- "drive item"
        private$remote <- !is.null(properties$remoteItem)
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
        children <- if(private$remote)
            self$properties$remoteItem$folder$childCount
        else self$properties$folder$childCount
        !is.null(children) && !is.na(children)
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

    list_items=function(path="", info=c("partial", "name", "all"), full_names=FALSE, filter=NULL, n=Inf, pagesize=1000)
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

        fullpath <- private$make_absolute_path(path)
        # possible fullpath formats -> string to append:
        # drives/xxx/root -> /children
        # drives/xxx/root:/foo/bar -> :/children
        # drives/xxx/items/yyy -> /children
        # drives/xxx/items/yyy:/foo/bar -> :/children
        op <- if(grepl(":/", fullpath)) paste0(fullpath, ":/children") else paste0(fullpath, "/children")
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

    get_parent_folder=function()
    {
        private$assert_is_not_remote()
        op <- private$make_absolute_path("..", FALSE)
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

    upload=function(src, dest=basename(src), blocksize=32768000, recursive=FALSE, parallel=FALSE)
    {
        private$assert_is_folder()

        # check if uploading a folder
        if(is.character(src) && dir.exists(src))
        {
            files <- dir(src, all.files=TRUE, no..=TRUE, recursive=recursive, full.names=FALSE)

            # dir() will always include subdirs if recursive is FALSE, must use horrible hack
            if(!recursive)
                files <- setdiff(files, list.dirs(src, recursive=FALSE, full.names=FALSE))

            # parallel can be:
            # - number: create cluster with this many workers
            # - cluster obj: use it
            # - TRUE: create cluster with 5 workers
            # - FALSE: serial
            if(isTRUE(parallel))
                parallel <- 5
            if(is.numeric(parallel))
            {
                parallel <- parallel::makeCluster(parallel)
                on.exit(parallel::stopCluster(parallel))
            }

            if(inherits(parallel, "cluster"))
            {
                parallel::parLapply(parallel, files, function(f, item, src, dest, blocksize)
                {
                    srcf <- file.path(src, f)
                    destf <- file.path(dest, f)
                    item$upload(srcf, destf, blocksize=blocksize)
                }, item=self, src=normalizePath(src), dest=dest, blocksize=blocksize)
            }
            else if(isFALSE(parallel))
            {
                for(f in files)
                {
                    srcf <- file.path(src, f)
                    destf <- file.path(dest, f)
                    private$upload_file(normalizePath(srcf), destf, blocksize=blocksize)
                }
            }
            else stop("Unknown value for 'parallel' argument", call.=FALSE)

            invisible(self$get_item(dest))
        }
        else private$upload_file(src, dest, blocksize)
    },

    download=function(dest=self$properties$name, overwrite=FALSE, recursive=FALSE, parallel=FALSE)
    {
        if(self$is_folder())
        {
            children <- self$list_items()
            isdir <- children$isdir

            if(!is.character(dest))
                stop("Must supply a destination folder", call.=FALSE)

            dest <- normalizePath(dest, mustWork=FALSE)
            dir.create(dest, showWarnings=FALSE)

            # parallel can be:
            # - number: create cluster with this many workers
            # - cluster obj: use it
            # - TRUE: create cluster with 5 workers
            # - FALSE: serial
            if(isTRUE(parallel))
                parallel <- 5
            if(is.numeric(parallel))
            {
                parallel <- parallel::makeCluster(parallel)
                on.exit(parallel::stopCluster(parallel))
            }

            if(inherits(parallel, "cluster"))
            {
                files <- children$name[!isdir]
                dirs <- children$name[isdir]

                # parallelise file downloads
                parallel::parLapply(parallel, files, function(f, item, dest, overwrite)
                {
                    item$get_item(f)$download(file.path(dest, f), overwrite=overwrite)
                }, item=self, dest=dest, overwrite=overwrite)

                # recursive call is done serially
                if(recursive) for(d in dirs)
                    self$get_item(d)$download(file.path(dest, d), overwrite=overwrite,
                                              parallel=parallel)
            }
            else if(isFALSE(parallel))
            {
                if(!recursive)
                    children <- children[!isdir, , drop=FALSE]
                for(f in children$name)
                    self$get_item(f)$download(file.path(dest, f), overwrite=overwrite,
                                              recursive=recursive, parallel=parallel)
            }
            else stop("Unknown value for 'parallel' argument", call.=FALSE)
        }
        else private$download_file(dest, overwrite)
    },

    load_dataframe=function(delim=NULL, ...)
    {
        private$assert_is_file()
        ext <- tolower(tools::file_ext(self$properties$name))
        if(is.null(delim))
        {
            delim <- if(ext == "csv") "," else if(ext == "csv2") ";" else "\t"
        }
        dat <- self$download(NULL)
        if(requireNamespace("readr"))
        {
            con <- rawConnection(dat, "r")
            on.exit(try(close(con), silent=TRUE))
            readr::read_delim(con, delim=delim, ...)
        }
        else utils::read.delim(text=rawToChar(dat), sep=delim, ...)
    },

    load_rdata=function(envir=parent.frame())
    {
        private$assert_is_file()
        private$assert_file_extension_is("rdata", "rda")
        rdata <- self$download(NULL)
        load(rawConnection(rdata, open="rb"), envir=envir)
    },

    load_rds=function()
    {
        private$assert_is_file()
        private$assert_file_extension_is("rds")
        rds <- self$download(NULL)
        unserialize(memDecompress(rds))
    },

    save_dataframe=function(df, file, delim=",", ...)
    {
        private$assert_is_folder()
        conn <- rawConnection(raw(0), open="r+b")
        if(requireNamespace("readr"))
            readr::write_delim(df, conn, delim=delim, ...)
        else utils::write.table(df, conn, sep=delim, ...)
        seek(conn, 0)
        self$upload(conn, file)
    },

    save_rdata=function(..., file, envir=parent.frame())
    {
        private$assert_is_folder()
        # save to a temporary file as saving to a connection disables compression
        tmpsave <- tempfile(fileext=".rdata")
        on.exit(unlink(tmpsave))
        save(..., file=tmpsave, envir=envir)
        self$upload(tmpsave, file)
    },

    save_rds=function(object, file)
    {
        private$assert_is_folder()
        # save to a temporary file to avoid dealing with memCompress/memDecompress hassles
        tmpsave <- tempfile(fileext=".rdata")
        on.exit(unlink(tmpsave))
        saveRDS(object, tmpsave)
        self$upload(tmpsave, file)
    },

    copy=function(dest, dest_folder_item=NULL)
    {
        path <- dirname(dest)
        body <- list(name=basename(dest))

        if(!is.null(dest_folder_item))
        {
            if(path != ".")
                warning("Destination folder object supplied; path will be ignored")
            body$parentReference <- list(
                driveId=dest_folder_item$properties$parentReference$driveId,
                id=dest_folder_item$properties$id
            )
        }
        else if(path != ".")
        {
            dest_folder_item <- private$get_drive()$get_item(path)
            body$parentReference <- list(
                driveId=dest_folder_item$properties$parentReference$driveId,
                id=dest_folder_item$properties$id
            )
        }

        self$do_operation("copy", body=body, http_verb="POST")
        invisible(NULL)
    },

    move=function(dest, dest_folder_item=NULL)
    {
        path <- dirname(dest)
        body <- list(name=basename(dest))

        if(!is.null(dest_folder_item))
        {
            if(path != ".")
                warning("Destination folder object supplied; path will be ignored")
            body$parentReference <- list(
                id=dest_folder_item$properties$id
            )
        }
        else if(path != ".")
        {
            dest_folder_item <- private$get_drive()$get_item(path)
            body$parentReference <- list(
                id=dest_folder_item$properties$id
            )
        }

        self$properties <- self$do_operation(body=body, encode="json", http_verb="PATCH")
        invisible(self)
    },

    get_path=function()
    {
        private$assert_is_not_remote()
        path <- private$make_absolute_path(use_itemid=FALSE)
        sub("^.+root:?/?", "/", path)
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

    # flag: whether this object is a shared file/folder on another drive
    # not actually needed! retained for backcompat
    remote=NULL,

    upload_file=function(src, dest, blocksize)
    {
        src <- normalize_src(src)
        on.exit(close(src$con))

        fullpath <- private$make_absolute_path(dest)

        # handle zero-length files correctly: cannot use resumable upload
        if(src$size == 0)
        {
            # possible fullpath formats -> string to append:
            # drives/xxx/root -> /content
            # drives/xxx/root:/foo/bar -> :/content
            # drives/xxx/items/yyy -> /content
            # drives/xxx/items/yyy:/foo/bar -> :/content
            op <- if(grepl(":/", fullpath))
                paste0(fullpath, ":/content")
            else paste0(fullpath, "/content")

            res <- call_graph_endpoint(self$token, op, http_verb="PUT", body=raw(0))
            return(invisible(ms_drive_item$new(self$token, self$tenant, httr::content(res))))
        }

        # possible fullpath formats -> string to append:
        # drives/xxx/root -> /createUploadSession
        # drives/xxx/root:/foo/bar -> :/createUploadSession
        # drives/xxx/items/yyy -> /createUploadSession
        # drives/xxx/items/yyy:/foo/bar -> :/createUploadSession
        op <- if(grepl(":/", fullpath))
            paste0(fullpath, ":/createUploadSession")
        else paste0(fullpath, "/createUploadSession")
        upload_dest <- call_graph_endpoint(self$token, op, http_verb="POST")$uploadUrl

        size <- src$size
        next_blockstart <- 0
        next_blockend <- size - 1
        repeat
        {
            next_blocksize <- min(next_blockend - next_blockstart + 1, blocksize)
            seek(src$con, next_blockstart)
            body <- readBin(src$con, "raw", next_blocksize)
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

    download_file=function(dest, overwrite)
    {
        private$assert_is_file()

        # TODO: make less hacky
        config <- if(is.character(dest))
            httr::write_disk(dest, overwrite=overwrite)
        else list()

        res <- self$do_operation("content", config=config, http_status_handler="pass")
        if(httr::status_code(res) >= 300)
        {
            if(is.character(dest))
                on.exit(file.remove(dest))
            httr::stop_for_status(res, paste0("complete operation. Message:\n",
                sub("\\.$", "", error_message(httr::content(res)))))
        }

        if(is.character(dest)) invisible(NULL) else httr::content(res, as="raw")
    },

    # dest = . or '' --> this item
    # dest = .. --> parent folder
    # dest = (childname) --> path to named child
    make_absolute_path=function(dest=".", use_itemid=getOption("microsoft365r_use_itemid_in_path"))
    {
        if(use_itemid == "remote")
            use_itemid <- !is.null(private$remoteItem)

        # use remote item props if present
        props <- if(!is.null(self$properties$remoteItem))
            self$properties$remoteItem
        else self$properties

        if(use_itemid)
            private$make_absolute_path_with_itemid(props, dest)
        else private$make_absolute_path_from_root(props, dest)
    },

    make_absolute_path_from_root=function(props, dest=".")
    {
        if(dest == ".")
            dest <- ""

        parent <- props$parentReference
        name <- props$name
        op <- if(name == "root")
            file.path("drives", parent$driveId, "root:")
        else
        {
            # null path means parent is the root folder
            if(is.null(parent$path))
                parent$path <- sprintf("/drives/%s/root:", parent$driveId)
            if(dest != "..")
                file.path(parent$path, name)
            else parent$path
        }
        if(dest != "..")
            op <- file.path(op, dest)
        utils::URLencode(enc2utf8(sub(":?/?$", "", op)))
    },

    # construct path using this item's ID
    # ".." not supported
    make_absolute_path_with_itemid=function(props, dest=".")
    {
        driveid <- props$parentReference$driveId
        id <- props$id
        base <- sprintf("drives/%s/items/%s", driveid, id)

        if(dest == "." || dest == "")
            return(base)
        else if(dest == "..")
            stop("Path with item ID to parent folder not supported", call.=FALSE)
        else if(substr(dest, 1, 1) == "/")
            stop("Absolute path incompatible with path starting from item ID", call.=FALSE)

        op <- sprintf("%s:/%s", base, dest)
        utils::URLencode(enc2utf8(op))
    },

    get_drive=function()
    {
        dummy_props <- list(id=self$properties$parentReference$driveId)
        ms_drive$new(self$token, self$tenant, dummy_props)
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
    },

    assert_is_not_remote=function()
    {
        if(!is.null(self$properties$remoteItem))
            stop("This method is not applicable for a remote item", call.=FALSE)
    },

    assert_file_extension_is=function(...)
    {
        ext <- tolower(tools::file_ext(self$properties$name))
        if(!(ext %in% unlist(list(...))))
            stop("Not an allowed file type")
    }
))


# alias for convenience
ms_drive_item$set("public", "list_files", overwrite=TRUE, ms_drive_item$public_methods$list_items)
