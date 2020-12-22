#' Personal OneDrive or SharePoint document library
#'
#' Class representing a personal OneDrive or SharePoint document library.
#'
#' @docType class
#' @section Fields:
#' - `token`: The token used to authenticate with the Graph host.
#' - `tenant`: The Azure Active Directory tenant for this drive.
#' - `type`: always "drive" for a drive object.
#' - `properties`: The drive properties.
#' @section Methods:
#' - `new(...)`: Initialize a new drive object. Do not call this directly; see 'Initialization' below.
#' - `delete(confirm=TRUE)`: Delete a drive. By default, ask for confirmation first.
#' - `update(...)`: Update the drive metadata in Microsoft Graph.
#' - `do_operation(...)`: Carry out an arbitrary operation on the drive.
#' - `sync_fields()`: Synchronise the R object with the drive metadata in Microsoft Graph.
#' - `list_items(path, info, full.names)`: List the files and folders under the specified path. See 'File and folder operations' below.
#' - `download_file(src, dest, overwrite)`: Download a file.
#' - `upload_file(src, dest, blocksize)`: Upload a file.
#' - `create_folder(path)`: Create a folder.
#' - `delete_item(path, confirm)`: Delete a file or folder.
#' - `get_item_properties(path)`: Get the properties (metadata) for a file or folder.
#' - `set_item_properties(path, ...)`: Set the properties for a file or folder.
#'
#' @section Initialization:
#' Creating new objects of this class should be done via the `get_drive` methods of the [ms_graph], [az_user] or [ms_site] classes. Calling the `new()` method for this class only constructs the R object; it does not call the Microsoft Graph API to retrieve or create the actual drive.
#'
#' @section File and folder operations:
#' This class exposes methods for carrying out common operations on files and folders.
#'
#' `list_items` lists the items under the specified path. It is the analogue of base R's `dir`/`list.files`. The arguments are
#' - `path`: The path.
#' - `info`: The information to return: either "partial", "name" or "all". If "partial", a data frame is returned containing the name, size and whether the item is a file or folder. If "name", a vector of file/folder names is returned. If "all", a data frame is returned containing _all_ the properties for each item (this can be large).
#' - `full.names`: Whether to prefix the full path to the names of the items.
#'
#' `download_file` and `upload_file` download and upload files from the local machine to the drive. For `upload_file`, the uploading is done in blocks of 32MB by default; you can change this by setting the `blocksize` argument. For technical reasons, the block size [must be a multiple of 320KB](https://docs.microsoft.com/en-us/graph/api/driveitem-createuploadsession?view=graph-rest-1.0#upload-bytes-to-the-upload-session).
#'
#' `create_folder` creates a folder with the specified path. Trying to create an already existing folder is an error.
#'
#' `delete_item` deletes a file or folder. By default, it will ask for confirmation first.
#'
#' `get_item_properties` returns an object of [ms_drive_item], containing the properties (metadata) for a given file or folder. The properties can be found in the `properties` field of this object.
#'
#' `set_item_properties` sets the properties (metadata) of a file or folder. The new properties should be specified as individual named arguments to the method. Any existing properties that aren't listed as arguments will retain their previous values or be recalculated based on changes to other properties, as appropriate.
#'
#' @seealso
#' [ms_graph], [ms_site], [ms_drive_item]
#'
#' [Microsoft Graph overview](https://docs.microsoft.com/en-us/graph/overview),
#' [REST API reference](https://docs.microsoft.com/en-us/graph/api/overview?view=graph-rest-1.0)
#'
#' @examples
#' \dontrun{
#'
#' gr <- get_graph_login()
#'
#' # shared document library for a SharePoint site
#' site <- gr$get_sharepoint_site("https://contoso.sharepoint.com/sites/O365-UG-123456")
#' drv <- site$get_drive()
#'
#' # personal OneDrive
#' gr2 <- get_graph_login("consumers")
#' me <- gr2$get_user()
#' mydrv <- me$get_drive()
#'
#' ## file/folder operationss
#' drv$list_items()
#' drv$list_items("path/to/folder", full.names=TRUE)
#'
#' # download a file -- default destination filename is taken from the source
#' drv$download_file("path/to/folder/data.csv")
#'
#' myfile <- drv$get_item_properties("myfile")
#' myfile$properties
#'
#' # rename a file
#' drv$set_item_properties("myfile", name="newname")
#'
#' }
#' @format An R6 object of class `ms_drive`, inheriting from `ms_object`.
#' @export
ms_drive <- R6::R6Class("ms_drive", inherit=ms_object,

public=list(

    initialize=function(token, tenant=NULL, properties=NULL)
    {
        self$type <- "drive"
        private$api_type <- "drives"
        super$initialize(token, tenant, properties)
    },

    list_items=function(path="/", info=c("partial", "name", "all"), full.names=FALSE, pagesize=1000)
    {
        info <- match.arg(info)
        opts <- switch(info,
            partial=list(`$select`="name,size,folder", `$top`=pagesize),
            name=list(`$select`="name", `$top`=pagesize),
            list(`$top`=pagesize)
        )

        children <- if(path != "/")
        {
            # get the item corresponding to this path, then list its children
            if(substr(path, 1, 1) != "/")
                path <- paste0("/", path)
            op <- file.path("root:", utils::URLencode(path, reserved=TRUE))
            item <- self$do_operation(op)
            self$do_operation(file.path("items", item$id, "children"), options=opts, simplify=TRUE)
        }
        else self$do_operation("root/children", options=opts, simplify=TRUE)

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

        if(full.names)
            df$name <- file.path(sub("^/", "", path), df$name)
        switch(info,
            partial=df[c("name", "size", "isdir")],
            name=df$name,
            all=
            {
                nms <- names(df)
                firstcols <- match(c("name", "size", "isdir"), nms)
                df[c(firstcols, setdiff(seq_along(nms), firstcols))]
            }
        )
    },

    download_file=function(src, dest=basename(src), overwrite=FALSE)
    {
        force(dest)
        src <- curl::curl_escape(sub("^/", "", src))
        path <- paste0("root:/", src, ":/content")
        res <- self$do_operation(path, config=httr::write_disk(dest, overwrite=overwrite))
        invisible(res)
    },

    upload_file=function(src, dest, blocksize=32768000)
    {
        dest <- curl::curl_escape(sub("^/", "", dest))
        path <- paste0("root:/", dest, ":/createUploadSession")

        con <- file(src, open="rb")
        on.exit(close(con))
        size <- file.size(src)

        upload_dest <- self$do_operation(path, http_verb="POST")$uploadUrl
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

    create_folder=function(path)
    {
        name <- basename(path)
        parent <- dirname(path)
        op <- if(parent %in% c(".", "/"))  # assume root
            "root/children"
        else paste0("root:/", sub("^/", "", parent), ":/children")
        parent <- curl::curl_escape(sub("^/", "", dirname(path)))
        body <- list(
            name=name,
            folder=named_list(),
            `@microsoft.graph.conflictBehavior`="fail"
        )
        res <- self$do_operation(op, body=body, http_verb="POST")
        ms_drive_item$new(self$token, self$tenant, res)
    },

    delete_item=function(path, confirm=TRUE)
    {
        self$get_file_properties(path)$delete(confirm=confirm)
    },

    get_item_properties=function(path)
    {
        op <- if(path == "/") "root" else file.path("root:", utils::URLencode(path))
        ms_drive_item$new(self$token, self$tenant, self$do_operation(op))
    },

    set_item_properties=function(path, ...)
    {
        op <- if(path == "/") "root" else file.path("root:", utils::URLencode(path))
        res <- self$do_operation(op, body=list(...), http_verb="PATCH")
        invisible(ms_drive_item$new(self$token, self$tenant, res))
    },

    print=function(...)
    {
        personal <- self$properties$driveType == "personal"
        name <- if(personal)
            paste0("<Personal OneDrive of ", self$properties$owner$user$displayName, ">\n")
        else paste0("<Document library '", self$properties$name, "'>\n")
        cat(name)
        cat("  directory id:", self$properties$id, "\n")
        if(!personal)
        {
            cat("  web link:", self$properties$webUrl, "\n")
            cat("  description:", self$properties$description, "\n")
        }
        cat("---\n")
        cat(format_public_methods(self))
        invisible(self)
    }
))


parse_upload_range <- function(response, blocksize)
{
    if(is.null(response$nextExpectedRanges))
        return(NULL)

    x <- as.numeric(strsplit(response$nextExpectedRanges[[1]], "-", fixed=TRUE)[[1]])
    if(length(x) == 1)
        x[2] <- x[1] + blocksize - 1
    x
}
