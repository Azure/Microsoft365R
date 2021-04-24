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
#' - `delete_item(path, confirm, by_item)`: Delete a file or folder. By default, ask for confirmation first. For personal OneDrive, deleting a folder will also automatically delete its contents; for business OneDrive or SharePoint document libraries, you may need to set `by_item=TRUE` to delete the contents first depending on your organisation's policies. Note that this can be slow for large folders.
#' - `get_item(path)`: Get an item representing a file or folder.
#' - `get_item_properties(path)`: Get the properties (metadata) for a file or folder.
#' - `set_item_properties(path, ...)`: Set the properties for a file or folder.
#'
#' @section Initialization:
#' Creating new objects of this class should be done via the `get_drive` methods of the [`ms_graph`], [`az_user`] or [`ms_site`] classes. Calling the `new()` method for this class only constructs the R object; it does not call the Microsoft Graph API to retrieve or create the actual drive.
#'
#' @section File and folder operations:
#' This class exposes methods for carrying out common operations on files and folders. They call down to the corresponding methods for the [`ms_drive_item`] class. In this context, any paths to child items are relative to the root folder of the drive.
#'
#' `open_item` opens the given file or folder in your browser. If the file has an unrecognised type, most browsers will attempt to download it.
#'
#' `list_items(path, info, full_names, pagesize)` lists the items under the specified path.
#'
#' `list_files` is a synonym for `list_items`.
#'
#' `download_file` and `upload_file` transfer files between the local machine and the drive. For `download_file`, the default destination folder is the current (working) directory of your R session. For `upload_file`, there is no default destination folder; make sure you specify the destination explicitly.
#'
#' `create_folder` creates a folder with the specified path. Trying to create an already existing folder is an error.
#'
#' `create_share_link(path, type, expiry, password, scope)` returns a shareable link to the item.
#'
#' `delete_item` deletes a file or folder. By default, it will ask for confirmation first.
#'
#' `get_item` retrieves the file or folder with the given path, as an  object of class [`ms_drive_item`].
#'
#' `get_item_properties` is a convenience function that returns the properties of a file or folder as a list.
#'
#' `set_item_properties` sets the properties of a file or folder. The new properties should be specified as individual named arguments to the method. Any existing properties that aren't listed as arguments will retain their previous values or be recalculated based on changes to other properties, as appropriate. You can also call the `update` method on the corresponding `ms_drive_item` object.
#'
#' @seealso
#' [`get_personal_onedrive`], [`get_business_onedrive`], [`ms_site`], [`ms_drive_item`]
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
#' # OneDrive for Business
#' busdrv <- get_business_onedrive("mycompany")
#'
#' # shared document library for a SharePoint site
#' site <- get_sharepoint_site("My site")
#' drv <- site$get_drive()
#'
#' ## file/folder operationss
#' drv$list_files()
#' drv$list_files("path/to/folder", full_names=TRUE)
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
        private$get_root()$list_items(path, info, full_names, pagesize)
    },

    upload_file=function(src, dest, blocksize=32768000)
    {
        private$get_root()$upload(src, dest, blocksize)
    },

    create_folder=function(path)
    {
        private$get_root()$create_folder(path)
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

    delete_item=function(path, confirm=TRUE, by_item=FALSE)
    {
        self$get_item(path)$delete(confirm=confirm, by_item=by_item)
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

    list_shared_items=function(info=c("partial", "name", "all"), full_names=FALSE, allow_external=TRUE, pagesize=1000)
    {
        info <- match.arg(info)
        opts <- switch(info,
            partial=list(`$select`="name,size,folder", `$top`=pagesize),
            name=list(`$select`="name", `$top`=pagesize),
            list(`$top`=pagesize)
        )
        if(allow_external)
            opts$allowExternal <- "true"
        children <- self$do_operation("sharedWithMe", options=opts, simplify=TRUE)

        # get remote file list as a data frame
        df <- private$get_paged_list(children, simplify=TRUE)$remoteItem

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
),

private=list(

    root=NULL,

    get_root=function()
    {
        if(is.null(private$root))
            private$root <- self$get_item("/")
        private$root
    }
))


# alias for convenience
ms_drive$set("public", "list_files", overwrite=TRUE, ms_drive$public_methods$list_items)


parse_upload_range <- function(response, blocksize)
{
    if(is_empty(response))
        return(NULL)

    # Outlook and Sharepoint/OneDrive teams not talking to each other....
    if(!is.null(response$NextExpectedRanges) && is.null(response$nextExpectedRanges))
        response$nextExpectedRanges <- response$NextExpectedRanges

    if(is.null(response$nextExpectedRanges))
        return(NULL)

    x <- as.numeric(strsplit(response$nextExpectedRanges[[1]], "-", fixed=TRUE)[[1]])
    if(length(x) == 1)
        x[2] <- x[1] + blocksize - 1
    x
}
