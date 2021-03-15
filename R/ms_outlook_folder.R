#' Outlook mail folder
#'
#' Class representing a folder in Outlook.
#'
#' @docType class
#' @section Fields:
#' - `token`: The token used to authenticate with the Graph host.
#' - `tenant`: The Azure Active Directory tenant for the parent drive.
#' - `type`: always "mail folder" for an Outlook folder object.
#' - `user_id`: the user ID of the Outlook account.
#' - `properties`: The item properties (metadata).
#' @section Methods:
#' - `new(...)`: Initialize a new object. Do not call this directly; see 'Initialization' below.
#' - `delete(confirm=TRUE)`: Delete this folder. By default, ask for confirmation first. Note that special folders cannot be deleted.
#' - `update(...)`: Update the item's properties (metadata) in Microsoft Graph.
#' - `do_operation(...)`: Carry out an arbitrary operation on the item.
#' - `sync_fields()`: Synchronise the R object with the item metadata in Microsoft Graph.
#' - `list_emails()`: List the emails in this folder.
#' - `get_email(message_id)`: Get the email with the specified ID.
#' - `create_email(...)`: Creates a new draft email in this folder, optionally sending it as well. See 'Sending emails' below.
#' - `delete_email(message_id, confim=TRUE)`: Deletes the specified email. By default, ask for confirmation first.
#' - `list_folders()`: List subfolders of this folder.
#' - `get_folder(folder_name, folder_id)`: Get a subfolder, either by the name or ID.
#' - `create_folder(folder_name)`: Create a new subfolder of this folder.
#' - `delete_folder(folder_name, folder_id, confirm=TRUE)`: Delete a subfolder. By default, ask for confirmation first.
#'
#' @section Initialization:
#' Creating new objects of this class should be done via the `get_folder`, `list_folders` or `create_folder` methods of this class or the [`ms_outlook`] class. Calling the `new()` method for this class only constructs the R object; it does not call the Microsoft Graph API to retrieve or create the actual folder.
#'
#' @section Creating and sending emails:
#' Outlook allows creating new draft emails in any folder, not just the Drafts folder (although that is the default for the Outlook app). To create a new email, call the `create_email()` method with the following arguments.
#'
#' @seealso
#' [`ms_outlook`], [`ms_outlook_email`]
#'
#' [Microsoft Graph overview](https://docs.microsoft.com/en-us/graph/overview),
#' [Outlook API reference](https://docs.microsoft.com/en-us/graph/api/resources/mail-api-overview?view=graph-rest-1.0)
#'
#' @examples
#' @format An R6 object of class `ms_outlook_folder`, inheriting from `ms_outlook_object`, which in turn inherits from `ms_object`.
#' @export
ms_outlook_folder <- R6::R6Class("ms_outlook_folder", inherit=ms_outlook_object,

public=list(

    user_id=NULL,

    initialize=function(token, tenant=NULL, properties=NULL, user_id=NULL)
    {
        if(is.null(user_id))
            stop("Must supply user ID", call.=FALSE)
        self$type <- "mail folder"
        self$user_id <- user_id
        private$api_type <- file.path("users", self$user_id, "mailFolders")
        super$initialize(token, tenant, properties)
    },

    list_emails=function(by=c("sent", "subject", "from", "importance"), order=c("descending", "ascending"), n=100)
    {
        opts <- list(`$orderby`=by)
        lst <- private$get_paged_list(self$do_operation("messages"), n=n)
        private$init_list_objects(lst, default_generator=ms_outlook_email, user_id=self$user_id)
    },

    get_email=function(message_id)
    {
        op <- file.path("messages", message_id)
        ms_outlook_email$new(self$token, self$tenant, self$do_operation(op), user_id=self$user_id)
    },

    create_email=function(body="", content_type=c("text", "html"), subject="", to=NULL, cc=NULL, bcc=NULL,
                          reply_to=NULL, attachments=NULL, send_now=FALSE)
    {
        content_type <- match.arg(content_type)
        req <- build_email_request(body, content_type, attachments, subject, to, cc, bcc)
        res <- ms_outlook_email$new(self$token, self$tenant,
            self$do_operation("messages", body=req, http_verb="POST"), user_id=self$user_id)

        if(send_now)
            res$send()
        res
    },

    delete_email=function(message_id, confirm=TRUE)
    {
        self$get_email(message_id)$delete(confirm=confirm)
    },

    list_folders=function()
    {
        lst <- private$get_paged_list(self$do_operation("childFolders"))
        private$init_list_objects(lst, default_generator=ms_outlook_folder, user_id=self$user_id)
    },

    get_folder=function(folder_name=NULL, folder_id=NULL)
    {
        assert_one_arg(folder_name, folder_id, msg="Supply exactly one of folder name or ID")

        if(!is.null(folder_id))
        {
            op <- file.path("users", self$user_id, "mailFolders", folder_id)
            res <- call_graph_endpoint(self$token, self$tenant, op)
            return(ms_outlook_folder$new(self$token, self$tenant, res, user_id=self$properties$id))
        }

        folders <- self$list_folders()
        wch <- which(sapply(folders, function(f) f$properties$displayName == folder_name))
        if(length(wch) != 1)
            stop("Invalid folder name '", folder_name, "'", call.=FALSE)
        else folders[[wch]]
    },

    create_folder=function(folder_name)
    {
        op <- file.path("childFolders", folder_name)
        ms_outlook_folder$new(self$token, self$tenant, self$do_operation(op, http_verb="POST"), user_id=self$user_id)
    },

    delete_folder=function(folder_name=NULL, folder_id=NULL, confirm=TRUE)
    {
        self$get_folder(folder_name, folder_id)$delete(confirm=confirm)
    },

    print=function(...)
    {
        cat("<Outlook folder '", self$properties$displayName, "'>\n", sep="")
        cat("  directory id:", self$properties$id, "\n")
        cat("---\n")
        cat(format_public_methods(self))
        invisible(self)
    }
))
