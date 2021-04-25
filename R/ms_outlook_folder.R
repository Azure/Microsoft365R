#' Outlook mail folder
#'
#' Class representing a folder in Outlook.
#'
#' @docType class
#' @section Fields:
#' - `token`: The token used to authenticate with the Graph host.
#' - `tenant`: The Azure Active Directory tenant for the email account.
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
#' - `create_email(...)`: Creates a new draft email in this folder, optionally sending it as well. See 'Creating and sending emails'.
#' - `delete_email(message_id, confim=TRUE)`: Deletes the specified email. By default, ask for confirmation first.
#' - `list_folders()`: List subfolders of this folder.
#' - `get_folder(folder_name, folder_id)`: Get a subfolder, either by the name or ID.
#' - `create_folder(folder_name)`: Create a new subfolder of this folder.
#' - `delete_folder(folder_name, folder_id, confirm=TRUE)`: Delete a subfolder. By default, ask for confirmation first.
#' - `copy(dest),move(dest)`: Copies or moves this folder to another folder. All the contents of the folder will also be copied/moved. The destination should be an object of class `ms_outlook_folder`.
#'
#' @section Initialization:
#' Creating new objects of this class should be done via the `get_folder`, `list_folders` or `create_folder` methods of this class or the [`ms_outlook`] class. Calling the `new()` method for this class only constructs the R object; it does not call the Microsoft Graph API to retrieve or create the actual folder.
#'
#' @section Creating and sending emails:
#' Outlook allows creating new draft emails in any folder, not just the Drafts folder (although that is the default location for the Outlook app, and the `ms_outlook` client class). To create a new email, call the `create_email()` method, which has the following signature:
#'```
#' create_email(body = "", content_type = c("text", "html"), subject = "",
#'              to = NULL, cc = NULL, bcc = NULL, reply_to = NULL, send_now = FALSE)
#' ```
#' - `body`: The body of the message. This should be a string or vector of strings, which will be pasted together with newlines as separators. You can also supply a message object as created by the blastula or emayili packages---see the examples below.
#' - `content_type`: The format of the body, either "text" (the default) or HTML.
#' - `subject`: The subject of the message.
#' - `to,cc,bcc,reply_to`: These should be lists of email addresses, in standard "user@host" format. You can also supply objects of class [`AzureGraph::az_user`] representing user accounts in Azure Active Directory.
#' - `send_now`: Whether the email should be sent immediately, or saved as a draft. You can send a draft email later with its `send()` method.
#'
#' This returns an object of class [`ms_outlook_email`], which has methods for making further edits and attaching files.
#'
#' You can also supply message objects as created by the blastula and emayili packages in the `body` argument. Note that blastula objects include attachments (if any), and emayili objects include attachments, recipients, and subject line; the corresponding arguments to `create_email()` will not be used in this case.
#'
#' To reply to or forward an email, first retrieve it using `get_email()` or `list_emails()`, and then call its `create_reply()`, `create_reply_all()` or `create_forward()` methods.
#'
#' @section Listing emails:
#' To list the emails in a folder, call the `list_emails()` method. This returns a list of objects of class [`ms_outlook_email`], and has the following signature:
#' ```
#' list_emails(by = "received desc", n = 100, pagesize = 10)
#' ```
#' - `by`: The sorting order of the message list. The possible fields are "received" (received date, the default), "from" and "subject". To sort in descending order, add a " desc". You can specify multiple sorting fields, with later fields used to break ties in earlier ones. The last sorting field is always "received desc" unless it appears earlier.
#' - `n`: The total number of emails to retrieve. The default is 100.
#' - `pagesize`: The number of emails per page. You can change this to a larger number to increase throughput, at the risk of running into timeouts.
#'
#' This returns a list of objects of class [`ms_outlook_email`].
#'
#' @seealso
#' [`ms_outlook`], [`ms_outlook_email`]
#'
#' [Microsoft Graph overview](https://docs.microsoft.com/en-us/graph/overview),
#' [Outlook API reference](https://docs.microsoft.com/en-us/graph/api/resources/mail-api-overview?view=graph-rest-1.0)
#'
#' @examples
#' \dontrun{
#'
#' outl <- get_personal_outlook()
#'
#' folder <- outl$get_folder("My folder")
#'
#' ##
#' ## listing emails
#' ##
#'
#' # the default: 100 most recent messages
#' folder$list_emails()
#'
#' # sorted by subject, then by most recent received date
#' folder$list_emails(by="subject")
#'
#' # sorted by from name in descending order, then by most recent received date
#' folder$list_emails(by="from desc")
#'
#' # retrieve a specific email:
#' # note the Outlook ID is NOT the same as the Internet message-id
#' email_id <- folder$list_emails()[[1]]$properties$id
#' folder$get_email(email_id)
#'
#' ##
#' ## creating/sending emails
#' ##
#'
#' # a simple text email with just a body:
#' # you can add other properties by calling the returned object's methods
#' folder$create_email("Hello from R")
#'
#' # HTML-formatted email with all necessary fields, sent immediately
#' folder$create_email("<emph>Emphatic hello</emph> from R",
#'     content_type="html",
#'     to="user@example.com",
#'     subject="example email",
#'     send_now=TRUE)
#'
#' # using blastula to create a HTML email with Markdown
#' bl_msg <- blastula::compose_email(md(
#' "
#' ## Hello!
#'
#' This is an email message that was generated by the blastula package.
#'
#' We can use **Markdown** formatting with the `md()` function.
#'
#' Cheers,
#'
#' The blastula team
#' "),
#'     footer=md("Sent via Microsoft365R"))
#' folder$create_email(bl_msg, subject="example blastula email")
#'
#'
#' # using emayili to create an email with attachments
#' ey_email <- emayili::envelope(
#'     text="Hello from emayili",
#'     to="user@example.com",
#'     subject="example emayili email") %>%
#'     emayili::attachment("mydocument.docx") %>%
#'     emayili::attachment("mydata.xlsx")
#' folder$create_email(ey_email)
#'
#' }
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

    list_emails=function(by="received desc", search=NULL, n=100, pagesize=10)
    {
        if(is.null(search))
        {
            order_by <- email_list_order(by)
            opts <- list(`$orderby`=order_by, `$top`=pagesize)
        }
        else
        {
            if(substr(search, 1, 1) != "" && substr(search, nchar(search), nchar(search)) != "")
                search <- paste0('"', search, '"')
            opts <- list(`$search`=search, `$top`=pagesize)
        }
        lst <- private$get_paged_list(self$do_operation("messages", options=opts), n=n)
        private$init_list_objects(lst, default_generator=ms_outlook_email, user_id=self$user_id)
    },

    get_email=function(message_id)
    {
        op <- file.path("messages", message_id)
        ms_outlook_email$new(self$token, self$tenant, self$do_operation(op), user_id=self$user_id)
    },

    create_email=function(body="", content_type=c("text", "html"), subject="", to=NULL, cc=NULL, bcc=NULL,
                          reply_to=NULL, send_now=FALSE)
    {
        content_type <- match.arg(content_type)
        req <- build_email_request(body, content_type, subject, to, cc, bcc, reply_to)
        res <- ms_outlook_email$new(self$token, self$tenant,
            self$do_operation("messages", body=req, http_verb="POST"), user_id=self$user_id)

        # must do this separately because large attachments require a valid message ID
        add_external_attachments(body, res)

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
        res <- self$do_operation("childFolders", body=list(displayName=folder_name), http_verb="POST")
        ms_outlook_folder$new(self$token, self$tenant, res, user_id=self$user_id)
    },

    delete_folder=function(folder_name=NULL, folder_id=NULL, confirm=TRUE)
    {
        self$get_folder(folder_name, folder_id)$delete(confirm=confirm)
    },

    copy=function(dest)
    {
        if(!inherits(dest, "ms_outlook_folder"))
            stop("Destination must be a folder object", call.=FALSE)

        body <- list(destinationId=dest$properties$id)
        ms_outlook_email$new(self$token, self$tenant, self$do_operation("copy", body=body, http_verb="POST"),
            user_id=self$user_id)
    },

    move=function(dest)
    {
        if(!inherits(dest, "ms_outlook_folder"))
            stop("Destination must be a folder object", call.=FALSE)

        on.exit(self$sync_fields())
        body <- list(destinationId=dest$properties$id)
        ms_outlook_email$new(self$token, self$tenant, self$do_operation("move", body=body, http_verb="POST"),
            user_id=self$user_id)
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


email_list_order <- function(vars)
{
    varmap <- c(
        "received"="receivedDateTime",
        "subject"="subject",
        "from"="from/emailAddress/name",
        "received desc"="receivedDateTime desc",
        "subject desc"="subject desc",
        "from desc"="from/emailAddress/name desc"
    )
    if(!all(vars %in% names(varmap)))
        stop("Unknown ordering field", call.=FALSE)
    if(!any(grepl("received", vars)))
        vars <- c(vars, "received desc")
    paste(varmap[vars], collapse=",")
}
