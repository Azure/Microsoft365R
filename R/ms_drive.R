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
#' - `list_items(...), list_files(...)`: List the files and folders under the specified path. See 'File and folder operations' below.
#' - `download_file(src, dest, overwrite)`: Download a file.
#' - `upload_file(src, dest, blocksize)`: Upload a file.
#' - `create_folder(path)`: Create a folder.
#' - `open_item(path)`: Open a file or folder.
#' - `create_share_link(...)`: Create a shareable link for a file or folder.
#' - `delete_item(path, confirm)`: Delete a file or folder.
#' - `get_item(path)`: Get an item representing a file or folder.
#' - `get_item_properties(path)`: Get the properties (metadata) for a file or folder.
#' - `set_item_properties(path, ...)`: Set the properties for a file or folder.
#'
#' @section Initialization:
#' Creating new objects of this class should be done via the `get_drive` methods of the [ms_graph], [az_user] or [ms_site] classes. Calling the `new()` method for this class only constructs the R object; it does not call the Microsoft Graph API to retrieve or create the actual drive.
#'
#' @section File and folder operations:
#' This class exposes methods for carrying out common operations on files and folders.
#'
#' `list_items(path, info, full_names, pagesize)` lists the items under the specified path. It is the analogue of base R's `dir`/`list.files`. Its arguments are
#' - `path`: The path.
#' - `info`: The information to return: either "partial", "name" or "all". If "partial", a data frame is returned containing the name, size and whether the item is a file or folder. If "name", a vector of file/folder names is returned. If "all", a data frame is returned containing _all_ the properties for each item (this can be large).
#' - `full_names`: Whether to prefix the full path to the names of the items.
#' - `pagesize`: The number of results to return for each call to the REST endpoint. You can try reducing this argument below the default of 1000 if you are experiencing timeouts.
#'
#' `list_files` is a synonym for `list_items`.
#'
#' `download_file` and `upload_file` download and upload files from the local machine to the drive. For `upload_file`, the uploading is done in blocks of 32MB by default; you can change this by setting the `blocksize` argument. For technical reasons, the block size [must be a multiple of 320KB](https://docs.microsoft.com/en-us/graph/api/driveitem-createuploadsession?view=graph-rest-1.0#upload-bytes-to-the-upload-session).
#'
#' `create_folder` creates a folder with the specified path. Trying to create an already existing folder is an error.
#'
#' `open_item` opens the given file or folder in your browser.
#'
#' `create_share_link(path, type, expiry, password, scope)` returns a shareable link to the item. Its arguments are
#' - `path`: The path.
#' - `type`: Either "view" for a read-only link, "edit" for a read-write link, or "embed" for a link that can be embedded in a web page. The last one is only available for personal OneDrive.
#' - `expiry`: How long the link is valid for. The default is 7 days; you can set an alternative like "15 minutes", "24 hours", "2 weeks", "3 months", etc. To leave out the expiry date, set this to NULL.
#' - `password`: An optional password to protect the link.
#' - `scope`: Optionally the scope of the link, either "anonymous" or "organization". The latter allows only users in your AAD tenant to access the link, and is only available for OneDrive for Business or SharePoint.
#'
#' This function returns a URL to access the item, for `type="view"` or "`type=edit"`. For `type="embed"`, it returns a list with components `webUrl` containing the URL, and `webHtml` containing a HTML fragment to embed the link in an IFRAME. The default is a viewable link, expiring in 7 days.
#'
#' `delete_item` deletes a file or folder. By default, it will ask for confirmation first.
#'
#' `get_item` returns an object of class [ms_drive_item], containing the properties (metadata) for a given file or folder and methods for working with it.
#'
#' `get_item_properties` is a convenience function that returns the properties of a file or folder as a list.
#'
#' `set_item_properties` sets the properties of a file or folder. The new properties should be specified as individual named arguments to the method. Any existing properties that aren't listed as arguments will retain their previous values or be recalculated based on changes to other properties, as appropriate.
#'
#' @seealso
#' [personal_onedrive], [business_onedrive], [ms_site], [ms_drive_item]
#'
#' [Microsoft Graph overview](https://docs.microsoft.com/en-us/graph/overview),
#' [OneDrive API reference](https://docs.microsoft.com/en-us/graph/api/resources/onedrive?view=graph-rest-1.0)
#'
#' @examples
#' \dontrun{
#'
#' # personal OneDrive
#' mydrv <- personal_onedrive()
#'
#' # OneDrive for Business
#' busdrv <- business_onedrive("mycompany")
#'
#' # shared document library for a SharePoint site
#' site <- sharepoint_site("https://mycompany.sharepoint.com/sites/my-site-name")
#' drv <- site$get_drive()
#'
#' ## file/folder operationss
#' drv$list_items()
#' drv$list_items("path/to/folder", full_names=TRUE)
#'
#' # download a file -- default destination filename is taken from the source
#' drv$download_file("path/to/folder/data.csv")
#'
#' # shareable links
#' drv$create_share_link("myfile")
#' drv$create_share_link("myfile", type="edit", expiry="24 hours")
#' drv$create_share_link("myfile", password="Use-strong-passwords!")
#'
#' # file metadata (name, date created, etc)
#' drv$get_item_properties("myfile")
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

    list_items=function(path="/", info=c("partial", "name", "all"), full_names=FALSE, pagesize=1000)
    {
        info <- match.arg(info)
        opts <- switch(info,
            partial=list(`$select`="name,size,folder", `$top`=pagesize),
            name=list(`$select`="name", `$top`=pagesize),
            list(`$top`=pagesize)
        )

        children <- if(path != "/")
        {
            path <- curl::curl_escape(gsub("^/|/$", "", path)) # remove any leading and trailing slashes
            self$do_operation(paste0("root:/", path, ":/children"), options=opts, simplify=TRUE)
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
        body <- list(
            name=name,
            folder=named_list(),
            `@microsoft.graph.conflictBehavior`="fail"
        )
        res <- self$do_operation(op, body=body, http_verb="POST")
        ms_drive_item$new(self$token, self$tenant, res)
    },

    download_file=function(src, dest=basename(src), overwrite=FALSE)
    {
        self$get_item(src)$download(dest, overwrite=overwrite)
    },

    create_share_link=function(path, type=c("view", "edit", "embed"), expiry="7 days", password=NULL, scope=NULL)
    {
        type <- match.arg(type)
        self$get_item(path)$get_share_link(type, expiry, password, scope)
    },

    open_item=function(path)
    {
        self$get_item(path)$open()
    },

    delete_item=function(path, confirm=TRUE)
    {
        self$get_item(path)$delete(confirm=confirm)
    },

    get_item=function(path)
    {
        op <- if(path != "/")
        {
            path <- curl::curl_escape(gsub("^/|/$", "", path)) # remove any leading and trailing slashes
            file.path("root:", path)
        }
        else "root"
        ms_drive_item$new(self$token, self$tenant, self$do_operation(op))
    },

    get_item_properties=function(path)
    {
        self$get_item(path)$properties
    },

    set_item_properties=function(path, ...)
    {
        self$get_item(path)$update(...)
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


# alias for convenience
ms_drive$set("public", "list_files", overwrite=TRUE, ms_drive$public_methods$list_items)


parse_upload_range <- function(response, blocksize)
{
    if(is.null(response$nextExpectedRanges))
        return(NULL)

    x <- as.numeric(strsplit(response$nextExpectedRanges[[1]], "-", fixed=TRUE)[[1]])
    if(length(x) == 1)
        x[2] <- x[1] + blocksize - 1
    x
}
