#' Outlook mail client
#'
#' Class representing a user's Outlook email account.
#'
#' @docType class
#' @section Fields:
#' - `token`: The token used to authenticate with the Graph host.
#' - `tenant`: The Azure Active Directory tenant for the email account.
#' - `type`: always "Outlook account" for an Outlook email account.
#' - `properties`: The item properties (metadata).
#' @section Methods:
#' - `new(...)`: Initialize a new object. Do not call this directly; see 'Initialization' below.
#' - `update(...)`: Update the account's properties (metadata) in Microsoft Graph.
#' - `do_operation(...)`: Carry out an arbitrary operation on the account.
#' - `sync_fields()`: Synchronise the R object with the account metadata in Microsoft Graph.
#' - `create_email(...)`: Creates a new email in the Drafts folder, optionally sending it as well. See 'Creating and sending emails'.
#' - `list_inbox_emails(...)`: List the emails in the Inbox folder. See 'Listing emails'.
#' - `get_inbox(),get_drafts(),get_sent_items(),get_deleted_items()`: Gets the special folder of that name. These folders are created by Outlook and exist in every email account.
#' - `list_folders(filter=NULL, n=Inf)`: List all folders in this account.
#' - `get_folder(folder_name, folder_id)`: Get a folder, either by the name or ID.
#' - `create_folder(folder_name)`: Create a new folder.
#' - `delete_folder(folder_name, folder_id, confirm=TRUE)`: Delete a folder. By default, ask for confirmation first. Note that special folders cannot be deleted.
#'
#' @section Initialization:
#' Creating new objects of this class should be done via the `get_personal_outlook()` or `get_business_outlook()` functions, or the `get_outlook` method of the [`az_user`] class. Calling the `new()` method for this class only constructs the R object; it does not call the Microsoft Graph API to retrieve the account information.
#'
#' @section Creating and sending emails:
#' To create a new email, call the `create_email()` method. The default behaviour is to create a new draft email in the Drafts folder, which can then be edited further to add attachments, recipients etc; or the email can be sent immediately.
#'
#' The `create_email()` method has the following signature:
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
#' This returns an object of class [`ms_outlook_email`], which has methods for making further edits, attaching files, replying, forwarding, and (re-)sending.
#'
#' You can also supply message objects as created by the blastula and emayili packages in the `body` argument. Note that blastula objects include attachments (if any), and emayili objects include attachments, recipients, and subject line; the corresponding arguments to `create_email()` will not be used in this case.
#'
#' @section Listing emails:
#' To list the emails in the Inbox, call the `list_emails()` method. This returns a list of objects of class [`ms_outlook_email`], and has the following signature:
#' ```
#' list_emails(by = "received desc", n = 100, pagesize = 10)
#' ```
#' - `by`: The sorting order of the message list. The possible fields are "received" (received date, the default), "from" and "subject". To sort in descending order, add a " desc". You can specify multiple sorting fields, with later fields used to break ties in earlier ones. The last sorting field is always "received desc" unless it appears earlier.
#' - `filter, n`: See below.
#' - `pagesize`: The number of emails per page. You can change this to a larger number to increase throughput, at the risk of running into timeouts.
#'
#' @section List methods generally:
#' All `list_*` methods have `filter` and `n` arguments to limit the number of results. The former should be an [OData expression](https://learn.microsoft.com/en-us/graph/query-parameters#filter-parameter) as a string to filter the result set on. The latter should be a number setting the maximum number of (filtered) results to return. The default values are `filter=NULL` and `n=100` for listing emails, and `n=Inf` for listing folders. If `n=NULL`, the `ms_graph_pager` iterator object is returned instead to allow manual iteration over the results.
#'
#' Support in the underlying Graph API for OData queries is patchy. Not all endpoints that return lists of objects support filtering, and if they do, they may not allow all of the defined operators. If your filtering expression results in an error, you can carry out the operation without filtering and then filter the results on the client side.
#' @seealso
#' [`ms_outlook_folder`], [`ms_outlook_email`]
#'
#' [Microsoft Graph overview](https://learn.microsoft.com/en-us/graph/overview),
#' [Outlook API reference](https://learn.microsoft.com/en-us/graph/api/resources/mail-api-overview?view=graph-rest-1.0)
#'
#' @examples
#' \dontrun{
#'
#' outl <- get_personal_outlook()
#'
#' ##
#' ## listing emails and folders
#' ##
#'
#' # the default: 100 most recent messages in the inbox
#' outl$list_emails()
#'
#' # sorted by subject, then by most recent received date
#' outl$list_emails(by="subject")
#'
#' # retrieve a specific email:
#' # note the Outlook ID is NOT the same as the Internet message-id
#' email_id <- outl$list_emails()[[1]]$properties$id
#' outl$get_email(email_id)
#'
#' # all folders in this account (including nested folders)
#' outl$list_folders()
#'
#' # draft (unsent) emails
#' dr <- outl$get_drafts()
#' dr$list_emails()
#'
#' # sent emails
#' sent <- outl$get_sent_items()
#' sent$list_emails()
#'
#' ##
#' ## creating/sending emails
#' ##
#'
#' # a simple text email with just a body (can't be sent)
#' outl$create_email("Hello from R")
#'
#' # HTML-formatted email with all necessary fields, sent immediately
#' outl$create_email("<emph>Emphatic hello</emph> from R",
#'     content_type="html",
#'     to="user@example.com",
#'     subject="example email",
#'     send_now=TRUE)
#'
#' # you can also create a blank email object and call its methods to add content
#' outl$create_email()$
#'     set_body("<emph>Emphatic hello</emph> from R", content_type="html")$
#'     set_recipients(to="user@example.com")$
#'     set_subject("example email")$
#'     add_attachment("mydocument.docx")$
#'     send()
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
#' outl$create_email(bl_msg, subject="example blastula email")
#'
#'
#' # using emayili to create an email with attachments
#' ey_email <- emayili::envelope(
#'     text="Hello from emayili",
#'     to="user@example.com",
#'     subject="example emayili email") %>%
#'     emayili::attachment("mydocument.docx") %>%
#'     emayili::attachment("mydata.xlsx")
#' outl$create_email(ey_email)
#'
#' }
#' @format An R6 object of class `ms_outlook`, inheriting from `ms_outlook_object`, which in turn inherits from `ms_object`.
#' @export
ms_outlook <- R6::R6Class("ms_outlook", inherit=ms_outlook_object,

public=list(

    initialize=function(token, tenant=NULL, properties=NULL)
    {
        self$type <- "Outlook account"
        private$api_type <- "users"
        super$initialize(token, tenant, properties)
    },

    delete=function(...)
    {
        stop("Cannot delete this object", call.=FALSE)
    },

    list_folders=function(filter=NULL, n=Inf)
    {
        private$make_basic_list("mailFolders", filter, n, user_id=self$properties$id)
    },

    get_folder=function(folder_name=NULL, folder_id=NULL)
    {
        if(is.null(folder_name) && is.null(folder_id))
            folder_name <- "inbox"

        assert_one_arg(folder_name, folder_id, msg="Supply at most one of folder name or ID")

        if(!is.null(folder_id))
        {
            op <- file.path("mailFolders", folder_id)
            return(ms_outlook_folder$new(self$token, self$tenant, self$do_operation(op), user_id=self$properties$id))
        }

        if(folder_name %in% special_email_folders)
        {
            op <- file.path("mailFolders", folder_name)
            return(ms_outlook_folder$new(self$token, self$tenant, self$do_operation(op), user_id=self$properties$id))
        }

        folders <- self$list_folders(filter=sprintf("displayName eq '%s'", folder_name))
        if(length(folders) != 1)
            stop("Invalid folder name '", folder_name, "'", call.=FALSE)
        else folders[[1]]
    },

    create_folder=function(folder_name)
    {
        res <- self$do_operation("mailFolders", body=list(displayName=folder_name), http_verb="POST")
        ms_outlook_folder$new(self$token, self$tenant, res, user_id=self$properties$id)
    },

    delete_folder=function(folder_name=NULL, folder_id=NULL, confirm=TRUE)
    {
        self$get_folder(folder_name, folder_id)$delete(confirm=confirm)
    },

    get_inbox=function()
    {
        self$get_folder("inbox")
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

    list_emails=function(...)
    {
        # use a dummy inbox folder object
        ms_outlook_folder$new(self$token, self$tenant, list(id="inbox"), user_id=self$properties$id)$
            list_emails(...)
    },

    create_email=function(...)
    {
        # use a dummy drafts folder object
        ms_outlook_folder$new(self$token, self$tenant, list(id="drafts"), user_id=self$properties$id)$
            create_email(...)
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
