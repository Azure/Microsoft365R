#' @format An R6 object of class `ms_outlook`, inheriting from `ms_outlook_object`, which in turn inherits from `ms_object`.
#' @export
ms_outlook <- R6::R6Class("ms_outlook", inherit=ms_outlook_object,

public=list(

    initialize=function(token, tenant=NULL, properties=NULL)
    {
        self$type <- "user"
        private$api_type <- "users"
        super$initialize(token, tenant, properties)
    },

    delete=function(...)
    {
        stop("Cannot delete this object", call.=FALSE)
    },

    list_folders=function()
    {
        lst <- private$get_paged_list(self$do_operation("mailFolders"))
        private$init_list_objects(lst, default_generator=ms_outlook_folder)
    },

    get_folder=function(folder_name=NULL, folder_id=NULL)
    {
        if(is.null(folder_name) && is.null(folder_id))
            folder_name <- "inbox"

        assert_one_arg(folder_name, folder_id, msg="Supply at most one of folder name or ID")

        if(!is.null(folder_id))
        {
            op <- file.path("mailFolders", folder_id)
            return(ms_outlook_folder$new(self$token, self$tenant, self$do_operation(op)))
        }

        if(folder_name %in% special_email_folders)
        {
            op <- file.path("mailFolders", folder_name)
            return(ms_outlook_folder$new(self$token, self$tenant, self$do_operation(op)))
        }

        folders <- self$list_folders()
        wch <- which(sapply(folders, function(f) f$properties$displayName == folder_name))
        if(length(wch) != 1)
            stop("Invalid folder name '", folder_name, "'", call.=FALSE)
        else folders[[wch]]
    },

    create_folder=function(folder_name)
    {
        res <- self$do_operation("mailFolders", body=list(displayName=folder_name), http_verb="POST")
        ms_outlook_folder$new(self$token, self$tenant, res)
    },

    delete_folder=function(folder_name=NULL, folder_id=NULL, confirm=TRUE)
    {
        self$get_folder(folder_name, folder_id)$delete(confirm=confirm)
    },

    get_inbox=function()
    {
        self$get_folder("inbox")
    },

    get_outbox=function()
    {
        self$get_folder("outbox")
    },

    get_sent_items=function()
    {
        self$get_folder("sentitems")
    },

    get_drafts=function()
    {
        self$get_folder("drafts")
    },

    get_deleted_items=function()
    {
        self$get_folder("deleteditems")
    },

    create_email=function(body="", content_type=c("text", "html"), subject="", to=NULL, cc=NULL, bcc=NULL,
                          attachments=NULL)
    {
        # use a dummy drafts folder object
        ms_outlook_folder$new(self$token, self$tenant, list(id="drafts"))$
            create_email(body, content_type, subject, to, cc, bcc, attachments)
    },

    print=function(...)
    {
        email <- if(!is_empty(self$properties$mail))
            self$properties$mail
        else self$properties$userPrincipalName
        if(is_empty(email))
            email <- "unknown"
        cat("<Outlook client for '", self$properties$displayName, "'>\n", sep="")
        cat("  email address:", email, "\n")
        cat("---\n")
        cat(format_public_methods(self))
        invisible(self)
    }
))

# special folders: assumed to exist in every account
special_email_folders <- c("inbox", "drafts", "outbox", "sentitems", "deleteditems", "junkemail", "archive", "clutter")
