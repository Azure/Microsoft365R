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
#' - `download_file(src, srcid, dest, overwrite)`: Download a file.
#' - `download_folder(src, srcid, dest, overwrite, recursive, parallel)`: Download a folder.
#' - `upload_file(src, dest, blocksize)`: Upload a file.
#' - `upload_folder(src, dest, blocksize, recursive, parallel)`: Upload a folder.
#' - `create_folder(path)`: Create a folder.
#' - `open_item(path, itemid)`: Open a file or folder.
#' - `create_share_link(...)`: Create a shareable link for a file or folder.
#' - `delete_item(path, itemid, confirm, by_item)`: Delete a file or folder. By default, ask for confirmation first. For personal OneDrive, deleting a folder will also automatically delete its contents; for business OneDrive or SharePoint document libraries, you may need to set `by_item=TRUE` to delete the contents first depending on your organisation's policies. Note that this can be slow for large folders.
#' - `get_item(path, itemid)`: Get an item representing a file or folder.
#' - `get_item_properties(path, itemid)`: Get the properties (metadata) for a file or folder.
#' - `set_item_properties(path, itemid, ...)`: Set the properties for a file or folder.
#' - `copy_item(path, itemid, dest, dest_item_id)`: Copy a file or folder.
#' - `move_item(path, itemid, dest, dest_item_id)`: Move a file or folder.
#' - `list_shared_items(...), list_shared_files(...)`: List the drive items shared with you. See 'Shared items' below.
#' - `load_dataframe(path, itemid, ...)`: Download a delimited file and return its contents as a data frame. See 'Saving and loading data' below.
#' - `load_rds(path, itemid)`: Download a .rds file and return the saved object.
#' - `load_rdata(path, itemid)`: Load a .RData or .Rda file into the specified environment.
#' - `save_dataframe(df, file, ...)` Save a dataframe to a delimited file.
#' - `save_rds(object, file)`: Save an R object to a .rds file.
#' - `save_rdata(..., file)`: Save the specified objects to a .RData file.
#'
#' @section Initialization:
#' Creating new objects of this class should be done via the `get_drive` methods of the [`ms_graph`], [`az_user`] or [`ms_site`] classes. Calling the `new()` method for this class only constructs the R object; it does not call the Microsoft Graph API to retrieve or create the actual drive.
#'
#' @section File and folder operations:
#' This class exposes methods for carrying out common operations on files and folders. They call down to the corresponding methods for the [`ms_drive_item`] class. In most cases an item can be specified either by path or ID. The former is more user-friendly but subject to change if the file is moved or renamed; the latter is an opaque string but is immutable regardless of file operations.
#'
#' `get_item(path, itemid)` retrieves a file or folder, as an object of class [`ms_drive_item`]. Specify either the path or ID, not both.
#'
#' `open_item` opens the given file or folder in your browser. If the file has an unrecognised type, most browsers will attempt to download it.
#'
#' `delete_item` deletes a file or folder. By default, it will ask for confirmation first.
#'
#' `create_share_link(path, itemid, type, expiry, password, scope)` returns a shareable link to the item.
#'
#' `get_item_properties` is a convenience function that returns the properties of a file or folder as a list.
#'
#' `set_item_properties` sets the properties of a file or folder. The new properties should be specified as individual named arguments to the method. Any existing properties that aren't listed as arguments will retain their previous values or be recalculated based on changes to other properties, as appropriate. You can also call the `update` method on the corresponding `ms_drive_item` object.
#'
#' - `copy_item` and `move_item` can take the destination location as either a full pathname (in the `dest` argument), or a name plus a drive item object (in the `dest_folder_item` argument). If the latter is supplied, any path in `dest` is ignored with a warning. Note that copying is an _asynchronous_ operation, meaning the method returns before the copy is complete.
#'
#' For copying and moving, the destination folder must exist beforehand. When copying/moving a large number of files, it's much more efficient to supply the destination folder in the `dest_folder_item` argument rather than as a path.
#'
#' `list_items(path, info, full_names, pagesize)` lists the items under the specified path.
#'
#' `list_files` is a synonym for `list_items`.
#'
#' `download_file` and `upload_file` transfer files between the local machine and the drive. For `download_file`, the default destination folder is the current (working) directory of your R session. For `upload_file`, there is no default destination folder; make sure you specify the destination explicitly.
#'
#' `download_folder` and `upload_folder` transfer all the files in a folder. If `recursive` is TRUE, all subfolders will also be transferred recursively. The `parallel` argument can have the following values:
#' - TRUE: A cluster with 5 workers is created
#' - A number: A cluster with this many workers is created
#' - A cluster object, created via the parallel package
#' - FALSE: The transfer is done serially
#' Transferring files in parallel can result in substantial speedup for a large number of small files.
#'
#' `create_folder` creates a folder with the specified path. Trying to create an already existing folder is an error.
#'
#' @section Saving and loading data:
#' The following methods are provided to simplify the task of loading and saving datasets and R objects. They call down to the corresponding methods for the `ms_drive_item` class. The `load_*`` methods allow specifying the file to be loaded by either a path or item ID.
#' - `load_dataframe` downloads a delimited file and returns its contents as a data frame. The delimiter can be specified with the `delim` argument; if omitted, this is "," if the file extension is .csv, ";" if the file extension is .csv2, and a tab otherwise. If the readr package is installed, the `readr::read_delim` function is used to parse the file, otherwise `utils::read.delim` is used. You can supply other arguments to the parsing function via the `...` argument.
#' - `save_dataframe` is the inverse of `load_dataframe`: it uploads the given data frame to a folder item. Specify the delimiter with the `delim` argument. The `readr::write_delim` function is used to serialise the data if that package is installed, and `utils::write.table` otherwise.
#' - `load_rds` downloads a .rds file and returns its contents as an R object. It is analogous to the base `readRDS` function but for OneDrive/SharePoint drive items.
#' - `save_rds` uploads a given R object as a .rds file, analogously to `saveRDS`.
#' - `load_rdata` downloads a .RData or .Rda file and loads its contents into the given environment. It is analogous to the base `load` function but for OneDrive/SharePoint drive items.
#' - `save_rdata` uploads the given R objects as a .RData file, analogously to `save`.
#'
#' @section Shared items:
#' The `list_shared_items` method shows the files and folders that have been shared with you. This is a named list of drive items, that you can use to access the shared files/folders. The arguments are:
#' - `allow_external`: Whether to include items that were shared from outside tenants. The default is FALSE.
#' - `filter, n`: See 'List methods' below.
#' - `pagesize`: The number of results to return for each call to the REST endpoint. You can try reducing this argument below the default of 1000 if you are experiencing timeouts.
#' - `info`: Deprecated, will be ignored. In previous versions, controlled the return type of the method.
#'
#' `list_shared_files` is a synonym for `list_shared_items`.
#'
#' Because of how the Graph API handles access to shared items linked in the root, you cannot directly access subitems of shared folders via the drive `get_item` method, like this: `drv$get_item("shared_folder/path/to/file")`. Instead, get the item into its own object, and use its `get_item` method: `drv$get_item("shared_folder")$get_item("path/to/file")`.
#'
#' @section List methods:
#' All `list_*` methods have `filter` and `n` arguments to limit the number of results. The former should be an [OData expression](https://learn.microsoft.com/en-us/graph/query-parameters#filter-parameter) as a string to filter the result set on. The latter should be a number setting the maximum number of (filtered) results to return. The default values are `filter=NULL` and `n=Inf`. If `n=NULL`, the `ms_graph_pager` iterator object is returned instead to allow manual iteration over the results.
#'
#' Support in the underlying Graph API for OData queries is patchy. Not all endpoints that return lists of objects support filtering, and if they do, they may not allow all of the defined operators. If your filtering expression results in an error, you can carry out the operation without filtering and then filter the results on the client side.
#' @seealso
#' [`get_personal_onedrive`], [`get_business_onedrive`], [`ms_site`], [`ms_drive_item`]
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
#' # rename a file -- item ID remains the same, while name is changed
#' obj <- drv$get_item("myfile")
#' drv$set_item_properties("myfile", name="newname")
#'
#' # retrieve the renamed object by ID
#' id <- obj$properties$id
#' obj2 <- drv$get_item(itemid=id)
#' obj$properties$id == obj2$properties$id  # TRUE
#'
#' # saving and loading data
#' drv$save_dataframe(iris, "path/to/iris.csv")
#' iris2 <- drv$load_dataframe("path/to/iris.csv")
#' identical(iris, iris2)  # TRUE
#'
#' drv$save_rds(iris, "path/to/iris.rds")
#' iris3 <- drv$load_rds("path/to/iris.rds")
#' identical(iris, iris3)  # TRUE
#'
#' # accessing shared files
#' shared_df <- drv$list_shared_items()
#' shared_df$remoteItem[[1]]$open()
#' shared_items <- drv$list_shared_items(info="items")
#' shared_items[[1]]$open()
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

    list_items=function(path="/", ...)
    {
        private$get_root()$list_items(path, ...)
    },

    upload_file=function(src, dest, blocksize=32768000)
    {
        private$get_root()$upload(src, dest, blocksize)
    },

    upload_folder=function(src, dest, blocksize=32768000, recursive=FALSE, parallel=FALSE)
    {
        private$get_root()$upload(src, dest, blocksize=blocksize, recursive=recursive, parallel=parallel)
    },

    create_folder=function(path)
    {
        private$get_root()$create_folder(path)
    },

    download_file=function(src=NULL, srcid=NULL, dest=basename(src), overwrite=FALSE)
    {
        self$get_item(src, srcid)$download(dest, overwrite=overwrite)
    },

    download_folder=function(src=NULL, srcid=NULL, dest=basename(src), overwrite=FALSE, recursive=FALSE,
                             parallel=FALSE)
    {
        self$get_item(src, srcid)$
            download(dest, overwrite=overwrite, recursive=recursive, parallel=parallel)
    },

    create_share_link=function(path=NULL, itemid=NULL, type=c("view", "edit", "embed"), expiry="7 days",
                               password=NULL, scope=NULL)
    {
        type <- match.arg(type)
        self$get_item(path, itemid)$get_share_link(type, expiry, password, scope)
    },

    open_item=function(path=NULL, itemid=NULL)
    {
        self$get_item(path, itemid)$open()
    },

    delete_item=function(path=NULL, itemid=NULL, confirm=TRUE, by_item=FALSE)
    {
        self$get_item(path, itemid)$delete(confirm=confirm, by_item=by_item)
    },

    get_item=function(path=NULL, itemid=NULL)
    {
        assert_one_arg(path, itemid, msg="Must supply one of item path or ID")
        op <- if(!is.null(itemid))
            file.path("items", itemid)
        else if(path != "/")
        {
            path <- curl::curl_escape(gsub("^/|/$", "", path)) # remove any leading and trailing slashes
            file.path("root:", path)
        }
        else "root"
        ms_drive_item$new(self$token, self$tenant, self$do_operation(op))
    },

    get_item_properties=function(path=NULL, itemid=NULL)
    {
        self$get_item(path, itemid)$properties
    },

    set_item_properties=function(path=NULL, itemid=NULL, ...)
    {
        self$get_item(path, itemid)$update(...)
    },

    copy_item=function(path=NULL, itemid=NULL, dest, dest_folder_item=NULL)
    {
        self$get_item(path, itemid)$copy(dest, dest_folder_item)
    },

    move_item=function(path=NULL, itemid=NULL, dest, dest_folder_item=NULL)
    {
        self$get_item(path, itemid)$move(dest, dest_folder_item)
    },

    list_shared_items=function(allow_external=TRUE, filter=NULL, n=Inf, pagesize=1000, info=NULL)
    {
        if(!is.null(info) && info != "items")
            warning("Ignoring 'info' argument, returning a list of drive items")

        opts <- list(`$top`=pagesize)
        if(allow_external)
            opts$allowExternal <- "true"
        if(!is.null(filter))
            opts$`filter` <- filter
        children <- self$do_operation("sharedWithMe", options=opts, simplify=FALSE)

        # get file list as a data frame, or return the iterator immediately if n is NULL
        out <- extract_list_values(self$get_list_pager(children), n)
        names(out) <- sapply(out, function(obj) obj$properties$name)
        out
    },

    load_dataframe=function(path=NULL, itemid=NULL, ...)
    {
        self$get_item(path, itemid)$load_dataframe(...)
    },

    load_rdata=function(path=NULL, itemid=NULL, envir=parent.frame())
    {
        self$get_item(path, itemid)$load_rdata(envir=envir)
    },

    load_rds=function(path=NULL, itemid=NULL)
    {
        self$get_item(path, itemid)$load_rds()
    },

    save_dataframe=function(df, file, ...)
    {
        folder <- dirname(file)
        if(folder == ".") folder <- "/"
        self$get_item(folder)$save_dataframe(df, basename(file), ...)
    },

    save_rdata=function(..., file, envir=parent.frame())
    {
        folder <- dirname(file)
        if(folder == ".") folder <- "/"
        self$get_item(folder)$save_rdata(..., file=basename(file), envir=envir)
    },

    save_rds=function(object, file)
    {
        folder <- dirname(file)
        if(folder == ".") folder <- "/"
        self$get_item(folder)$save_rds(object, file=basename(file))
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


# aliases for convenience
ms_drive$set("public", "list_files", overwrite=TRUE, ms_drive$public_methods$list_items)

ms_drive$set("public", "list_shared_files", overwrite=TRUE, ms_drive$public_methods$list_shared_items)

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
