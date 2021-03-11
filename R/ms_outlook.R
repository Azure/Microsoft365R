#' @format An R6 object of class `ms_outlook`, inheriting from `az_user`.
#' @export
ms_outlook <- R6::R6Class("ms_outlook", inherit=az_user,

public=list(

    initialize=function(token, tenant=NULL, properties=NULL)
    {
        super$initialize(token, tenant, properties)
    },

    delete=function(...)
    {
        stop("Cannot delete this object", call.=FALSE)
    },

    list_folders=function()
    {
        lst <- private$get_paged_list(self$do_operation("mailFolders"))
        private$init_list_objects(lst, "ms_email_folder")
    },

    get_folder=function(folder_name=NULL, folder_id=NULL)
    {
        if(is.null(folder_name) && is.null(folder_id))
            folder_name <- "inbox"

        assert_one_arg(folder_name, folder_id, msg="Supply at most one of folder name or ID")

        if(!is.null(folder_id))
        {
            op <- file.path("mailFolders", folder_id)
            return(ms_email_folder$new(self$token, self$tenant, self$do_operation(op)))
        }

        # special folders
        specials <- c("inbox", "drafts", "outbox", "sentitems", "deleteditems", "junkemail", "archive", "clutter")
        if(folder_name %in% specials)
        {
            op <- file.path("mailFolders", folder_name)
            return(ms_email_folder$new(self$token, self$tenant, self$do_operation(op)))
        }

        folders <- self$list_folders()
        wch <- which(sapply(folders, function(f) f$properties$displayName == folder_name))
        if(length(wch) != 1)
            stop("Invalid folder name '", folder_name, "'", call.=FALSE)
        else folders[[wch]]
    },

    delete_folder=function(folder_name=NULL, folder_id=NULL, confirm=TRUE)
    {
        self$get_folder(folder_name, folder_id)$delete(confirm=confirm)
    },

    get_inbox=function()
    {
        self$get_folder("inbox")
    },

    create_email=function(body, attachments=NULL, to=NULL, cc=NULL, bcc=NULL)
    {

    }
))
