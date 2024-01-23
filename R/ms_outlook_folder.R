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
#' - `list_emails(...)`: List the emails in this folder.
#' - `get_email(message_id)`: Get the email with the specified ID.
#' - `create_email(...)`: Creates a new draft email in this folder, optionally sending it as well. See 'Creating and sending emails'.
#' - `delete_email(message_id, confim=TRUE)`: Deletes the specified email. By default, ask for confirmation first.
#' - `list_folders(filter=NULL, n=Inf)`: List subfolders of this folder.
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
#' list_emails(by = "received desc", search = NULL, filter = NULL, n = 100, pagesize = 10)
#' ```
#' - `by`: The sorting order of the message list. The possible fields are "received" (received date, the default), "from" and "subject". To sort in descending order, add a " desc". You can specify multiple sorting fields, with later fields used to break ties in earlier ones. The last sorting field is always "received desc" unless it appears earlier.
#' - `search`: An optional string to search for. Only emails that contain the search string will be returned. See the [description of this parameter](https://learn.microsoft.com/en-us/graph/query-parameters#search-parameter) for more information.
#' - `filter, n`: See below.
#' - `pagesize`: The number of emails per page. You can change this to a larger number to increase throughput, at the risk of running into timeouts.
#'
#' Currently, searching and filtering the message list is subject to some limitations. You can only specify one of `search` and `filter`; searching and filtering at the same time will not work. Ordering the results is only allowed if neither a search term nor a filtering expression is present. If searching or filtering is done, the result is always sorted by date.
#'
#' @section List methods generally:
#' All `list_*` methods have `filter` and `n` arguments to limit the number of results. The former should be an [OData expression](https://learn.microsoft.com/en-us/graph/query-parameters#filter-parameter) as a string to filter the result set on. The latter should be a number setting the maximum number of (filtered) results to return. The default values are `filter=NULL` and `n=100` for listing emails, and `n=Inf` for listing folders. If `n=NULL`, the `ms_graph_pager` iterator object is returned instead to allow manual iteration over the results.
#'
#' Support in the underlying Graph API for OData queries is patchy. Not all endpoints that return lists of objects support filtering, and if they do, they may not allow all of the defined operators. If your filtering expression results in an error, you can carry out the operation without filtering and then filter the results on the client side.
#' @seealso
#' [`ms_outlook`], [`ms_outlook_email`]
#'
#' [Microsoft Graph overview](https://learn.microsoft.com/en-us/graph/overview),
#' [Outlook API reference](https://learn.microsoft.com/en-us/graph/api/resources/mail-api-overview?view=graph-rest-1.0)
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
#' # searching the list
#' folder$list_emails(search="important information")
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

    list_emails=function(by="received desc", search=NULL, filter=NULL, n=100, pagesize=10)
    {
        # search term must have double quotes around it
        if(!is.null(search) && substr(search, 1, 1) != "" && substr(search, nchar(search), nchar(search)) != "")
            search <- paste0('"', search, '"')

        # by only works with no filter and no search
        order_by <- if(is.null(filter) && is.null(search)) email_list_order(by)

        opts <- list(`$orderby`=order_by, `$search`=search, `$filter`=filter, `$top`=pagesize)
        pager <- self$get_list_pager(self$do_operation("messages", options=opts), default_generator=ms_outlook_email,
                                     user_id=self$user_id)
        extract_list_values(pager, n)
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

    list_folders=function(filter=NULL, n=Inf)
    {
        private$make_basic_list("childFolders", filter, n, user_id=self$user_id)
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

        folders <- self$list_folders(filter=sprintf("displayName eq '%s'", folder_name))
        if(length(folders) != 1)
            stop("Invalid folder name '", folder_name, "'", call.=FALSE)
        else folders[[1]]
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
