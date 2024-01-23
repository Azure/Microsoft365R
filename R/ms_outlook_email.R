#' Outlook mail message
#'
#' Class representing an Outlook mail message. The one class represents both sent and unsent (draft) emails.
#'
#' @docType class
#' @section Fields:
#' - `token`: The token used to authenticate with the Graph host.
#' - `tenant`: The Azure Active Directory tenant for the email account.
#' - `type`: always "email" for an Outlook mail message.
#' - `properties`: The item properties (metadata).
#' @section Methods:
#' - `new(...)`: Initialize a new object. Do not call this directly; see 'Initialization' below.
#' - `delete(confirm=TRUE)`: Delete this email. By default, ask for confirmation first.
#' - `update(...)`: Update the email's properties (metadata) in Microsoft Graph.
#' - `do_operation(...)`: Carry out an arbitrary operation on the email.
#' - `sync_fields()`: Synchronise the R object with the email metadata in Microsoft Graph.
#' - `set_body(body=NULL, content_type=NULL)`: Update the email body. See 'Editing an email' below.
#' - `set_subject(subject)`: Update the email subject line.
#' - `set_recipients(to=NULL, cc=NULL, bcc=NULL)`: Set the recipients for the email, overwriting any existing recipients.
#' - `add_recipients(to=NULL, cc=NULL, bcc=NULL)`: Adds recipients for the email, leaving existing ones unchanged.
#' - `set_reply_to(reply_to=NULL)`: Sets the reply-to field for the email.
#' - `add_attachment(object, ...)`: Adds an attachment to the email. See 'Attachments' below.
#' - `add_image(object)`: Adds an inline image to the email.
#' - `get_attachment(attachment_name=NULL, attachment_id=NULL)`: Gets an attachment, either by name or ID. Note that attachments don't need to have unique names; if multiple attachments share the same name, the method throws an error.
#' - `list_attachments(filter=NULL, n=Inf)`: Lists the current attachments for the email.
#' - `remove_attachment(attachment_name=NULL, attachment_id=NULL, confirm=TRUE)`: Removes an attachment from the email. By default, ask for confirmation first.
#' - `download_attachment(attachment_name=NULL, attachment_id=NULL, ...)`: Downloads an attachment. This is only supported for file attachments (not URLs).
#' - `send()`: Sends an email.  See 'Sending, replying and forwarding'.
#' - `create_reply(comment="", send_now=FALSE)`: Replies to the sender of an email.
#' - `create_reply_all(comment="", send_now=FALSE)`: Replies to the sender and all recipients of an email.
#' - `create_forward(comment="", to=NULL, cc=NULL, bcc=NULL, send_now=FALSE)`: Forwards the email to other recipients.
#' - `copy(dest),move(dest)`: Copies or moves the email to the destination folder.
#' - `get_message_headers`: Retrieves the Internet message headers for an email, as a named character vector.
#'
#' @section Initialization:
#' Creating new objects of this class should be done via the the appropriate methods for the [`ms_outlook`] or [`ms_outlook_folder`] classes. Calling the `new()` method for this class only constructs the R object; it does not call the Microsoft Graph API to retrieve or create the actual folder.
#'
#' @section Editing an email:
#' This class exposes several methods for updating the properties of an email. They should work both for unsent (draft) emails and sent ones, although they make most sense in the context of editing drafts.
#'
#' `set_body(body, content_type)` updates the message body of the email. This has 2 arguments: `body` which is the body text itself, and `content_type` which should be either "text" or "html". For both arguments, you can set the value to NULL to leave the current property unchanged. The `body` argument can also be a message object from either the blastula or emayili packages, much like when creating a new email.
#'
#' `set_subject(subject)` sets the subject line of the email.
#'
#' `set_recipients(to, cc, bcc)` sets or clears the recipients of the email. The `to`, `cc` and `bcc` arguments should be lists of either email addresses as character strings, or objects of class `az_user` representing a user account in Azure Active Directory. The default behaviour is to overwrite any existing recipients; to avoid this, pass `NA` as the value for the relevant argument. Alternatively, you can use the `add_recipients()` method.
#'
#' `add_recipients(to, cc, bcc)` is like `set_recipients()` but leaves existing recipients unchanged.
#'
#' `set_reply_to(reply_to)` sets or clears the reply-to address for the email. Leave the `reply_to` argument at its default NULL value to clear this property.
#'
#' @section Attachments:
#' This class exposes the following methods for working with attachments.
#'
#' `add_attachment(object, type, expiry, password, scope)` adds an attachment to the email. The arguments are as follows:
#' - `object`: A character string containing a filename or URL, or an object of class [`ms_drive_item`] representing a file in OneDrive or SharePoint. In the latter case, a shareable link to the drive item will be attached to the email, with the link details given by the other arguments.
#' - `type, expiry, password, scope`: The specifics for the shareable link to attach to the email, if `object` is a drive item. See the `create_share_link()` method of the [`ms_drive_item`] class; the default is to create a read-only link valid for 7 days.
#'
#' `add_image(object)` adds an image as an _inline_ attachment, ie, as part of the message body. The `object` argument should be a filename, and the message content type will be set to "html" if it is not already. Currently Microsoft365R does minimal formatting of the image; consider using a package like blastula for more control over the layout of inline images.
#'
#' `list_attachments()` lists the attachments for the email, including inline images. This will be a list of objects of class [`ms_outlook_attachment`] containing the metadata for the attachments.
#'
#' `get_attachment(attachment_name, attachment_id)`: Retrieves the metadata for an attachment, as an object of class `ms_outlook_attachment`. Note that multiple attachments can share the same name; in this case, you must specify the ID of the attachment.
#'
#' `download_attachment(attachment_name, attachment_id, dest, overwrite)`: Downloads a file attachment. The default destination filename is the name of the attachment.
#'
#' `remove_attachment(attachment_name, attachment_id)` removes (deletes) an attachment.
#'
#' @section Sending, replying and forwarding:
#' Microsoft365R's default behaviour when creating, replying or forwarding emails is to create a draft message object, to allow for further edits. The draft is saved in the Drafts folder by default, and can be sent later by calling its `send()` method.
#'
#' The methods for replying and forwarding are `create_reply()`, `create_reply_all()` and `create_forward()`. The first argument to these is the reply text, which will appear above the current message text in the body of the reply. For `create_forward()`, the other arguments are `to`, `cc` and `bcc` to specify the recipients of the forwarded email.
#'
#' @section Other methods:
#' The `copy()` and `move()` methods copy and move an email to a different folder. The destination should be an object of class `ms_outlook_folder`.
#'
#' The `get_message_headers()` method retrieves the Internet message headers for the email, as a named character vector.
#'
#' @section List methods:
#' All `list_*` methods have `filter` and `n` arguments to limit the number of results. The former should be an [OData expression](https://learn.microsoft.com/en-us/graph/query-parameters#filter-parameter) as a string to filter the result set on. The latter should be a number setting the maximum number of (filtered) results to return. The default values are `filter=NULL` and `n=Inf`. If `n=NULL`, the `ms_graph_pager` iterator object is returned instead to allow manual iteration over the results.
#'
#' Support in the underlying Graph API for OData queries is patchy. Not all endpoints that return lists of objects support filtering, and if they do, they may not allow all of the defined operators. If your filtering expression results in an error, you can carry out the operation without filtering and then filter the results on the client side.
#' @seealso
#' [`ms_outlook`], [`ms_outlook_folder`], [`ms_outlook_attachment`]
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
#' ## creating a new email
#' ##
#'
#' # a blank text email
#' em <- outl$create_email()
#'
#' # add a body
#' em$set_body("Hello from R", content_type="html")
#'
#' # add recipients
#' em$set_recipients(to="user@example.com")
#'
#' # add subject line
#' em$set_subject("example email")
#'
#' # add an attachment
#' em$add_attachment("mydocument.docx")
#'
#' # add a shareable link to a file in OneDrive
#' mysheet <- get_personal_onedrive()$get_item("documents/mysheet.xlsx")
#' em$add_attachment(mysheet)
#'
#' # add an inline image
#' em$add_image("myggplot.jpg")
#'
#' # oops, wrong recipient, it should be someone else
#' # this removes user@example.com from the to: field
#' em$set_recipients(to="user2@example.com")
#'
#' # and we should also cc a third user
#' em$add_recipients(cc="user3@example.com")
#'
#' # send it
#' em$send()
#'
#' # you can also compose an email as a pipeline
#' outl$create_email()$
#'     set_body("Hello from R")$
#'     set_recipients(to="user2@example.com", cc="user3@example.com")$
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
#' outl$create_email()
#'     set_body(bl_msg)$
#'     set_subject("example blastula email")
#'
#'
#' ##
#' ## replying and forwarding
#' ##
#'
#' # get the most recent email in the Inbox
#' em <- outl$list_emails()[[1]]
#'
#' # reply to the message sender, cc'ing Carol
#' em$create_reply("I agree")$
#'     add_recipients(cc="carol@example.com")$
#'     send()
#'
#' # reply to everyone, setting the reply-to address
#' em$create_reply_all("Please do not reply")$
#'     set_reply_to("do_not_reply@example.com")$
#'     send()
#'
#' # forward to Dave
#' em$create_forward("FYI", to="dave@example.com")$
#'     send()
#'
#'
#' ##
#' ## attachments
#' ##
#'
#' # download an attachment by name (assumes there is only one 'myfile.docx')
#' em$download_attachment("myfile.docx")
#'
#' # a more reliable way: get the list of attachments, and download via the object
#' atts <- em$list_attachments()
#' atts[[1]]$download()
#'
#' # add and remove an attachment
#' em$add_attachment("anotherfile.pptx")
#' em$remove_attachment("anotherfile.pptx")
#'
#'
#' ##
#' ## moving and copying
#' ##
#'
#' # copy an email to a nested folder: /folder1/folder2
#' dest <- outl$get_folder("folder1")$get_folder("folder2")
#' em$copy(dest)
#'
#' # move it instead
#' em$move(dest)
#'
#' }
#' @format An R6 object of class `ms_outlook_email`, inheriting from `ms_outlook_object`, which in turn inherits from `ms_object`.
#' @export
ms_outlook_email <- R6::R6Class("ms_outlook_email", inherit=ms_outlook_object,

public=list(

    user_id=NULL,

    initialize=function(token, tenant=NULL, properties=NULL, user_id=NULL)
    {
        if(is.null(user_id))
            stop("Must supply user ID", call.=FALSE)
        self$type <- "email"
        self$user_id <- user_id
        private$api_type <- file.path("users", self$user_id, "messages")
        super$initialize(token, tenant, properties)
    },

    set_body=function(body=NULL, content_type=NULL)
    {
        if(is.null(body) && is.null(content_type))
            return(self)

        if(is.null(body))
            body <- self$properties$body$content
        if(is.null(content_type))
            content_type <- self$properties$body$contentType
        req <- build_email_request(body, content_type)
        do.call(self$update, req)
    },

    set_subject=function(subject)
    {
        self$update(subject=subject)
    },

    set_recipients=function(to=NULL, cc=NULL, bcc=NULL)
    {
        if(is_empty(to) && is_empty(cc) && is_empty(bcc))
            message("Clearing all recipients")
        do.call(self$update, build_email_recipients(to, cc, bcc, NA))
    },

    add_recipients=function(to=NULL, cc=NULL, bcc=NULL)
    {
        find_address <- function(x)
        {
            x$emailAddress$address
        }
        current_to <- lapply(self$properties$toRecipients, find_address)
        current_cc <- lapply(self$properties$ccRecipients, find_address)
        current_bcc <- lapply(self$properties$bccRecipients, find_address)

        self$set_recipients(c(current_to, to), c(current_cc, cc), c(current_bcc, bcc))
    },

    set_reply_to=function(reply_to=NULL)
    {
        if(is_empty(reply_to))
            message("Clearing reply-to")

        # possible bug: can only set reply-to if 1 other recipient field is included in request
        recipients <- build_email_recipients(NA, NA, NA, reply_to)
        recipients$toRecipients <- self$properties$toRecipients
        do.call(self$update, recipients)
    },

    add_attachment=function(object, type=c("view", "edit", "embed"), expiry="7 days", password=NULL, scope=NULL)
    {
        att <- private$make_attachment(object, FALSE, match.arg(type), expiry, password, scope)
        if(!is_empty(att))  # check for large attachment
            self$do_operation("attachments", body=att, http_verb="POST")
        self$sync_fields()
    },

    add_image=function(object)
    {
        if(self$properties$body$contentType != "html")
            warning("Message body will be converted to HTML", call.=FALSE)

        att <- private$make_attachment(object, TRUE)
        if(!is_empty(att))
            self$do_operation("attachments", body=att, http_verb="POST")
        body <- c(self$properties$body$content,
            sprintf('<img src="cid:%s"/>', att$name))
        self$set_body(body=body, content_type="html")
    },

    get_attachment=function(attachment_name=NULL, attachment_id=NULL)
    {
        assert_one_arg(attachment_name, attachment_id, msg="Supply exactly one of attachment name or ID")
        if(is.null(attachment_id))
        {
            # filter arg not working with attachments?
            atts <- self$list_attachments()
            att_names <- sapply(atts, function(a) a$properties$name)
            wch <- which(att_names == attachment_name)
            if(length(wch) == 0)
                stop("Attachment '", attachment_name, "' not found", call.=FALSE)
            if(length(wch) > 1)
                stop("More than one attachment named '", attachment_name, "'", call.=FALSE)
            return(atts[[wch]])
        }

        fields <- c("id", "name", "contentType", "size", "isInline", "lastModifiedDateTime")
        res <- self$do_operation(file.path("attachments", attachment_id),
            options=list(select=paste(fields, collapse=",")))
        ms_outlook_attachment$new(self$token, self$tenant, res,
            user_id=self$user_id, message_id=self$properties$id)
    },

    list_attachments=function(filter=NULL, n=Inf)
    {
        fields <- c("id", "name", "contentType", "size", "isInline", "lastModifiedDateTime")
        opts <- list(select=paste(fields, collapse=","))
        if(!is.null(filter))
            opts$`filter` <- filter
        pager <- self$get_list_pager(self$do_operation("attachments", options=opts),
            user_id=self$user_id, message_id=self$properties$id)
        extract_list_values(pager, n)
    },

    remove_attachment=function(attachment_name=NULL, attachment_id=NULL, confirm=TRUE)
    {
        self$get_attachment(attachment_name, attachment_id)$delete(confirm=confirm)
        self$sync_fields()
    },

    download_attachment=function(attachment_name=NULL, attachment_id=NULL, ...)
    {
        self$get_attachment(attachment_name, attachment_id)$download(...)
    },

    send=function()
    {
        self$do_operation("send", http_verb="POST")
        self$sync_fields()
    },

    create_reply=function(comment="", send_now=FALSE)
    {
        op <- "createReply"
        body <- list(comment=make_reply_comment(comment))
        reply <- ms_outlook_email$new(self$token, self$tenant,
            self$do_operation(op, body=body, http_verb="POST"), user_id=self$user_id)
        if(send_now)
            reply$send()
        reply
    },

    create_reply_all=function(comment="", send_now=FALSE)
    {
        op <- "createReplyAll"
        body <- list(comment=make_reply_comment(comment))
        reply <- ms_outlook_email$new(self$token, self$tenant,
            self$do_operation(op, body=body, http_verb="POST"), user_id=self$user_id)
        if(send_now)
            reply$send()
        reply
    },

    create_forward=function(comment="", to=NULL, cc=NULL, bcc=NULL, send_now=FALSE)
    {
        op <- "createforward"
        body <- list(
            comment=make_reply_comment(comment),
            message=build_email_recipients(to, cc, bcc, NA)
        )
        reply <- ms_outlook_email$new(self$token, self$tenant,
            self$do_operation(op, body=body, http_verb="POST"), user_id=self$user_id)
        if(send_now)
            reply$send()
        reply
    },

    get_message_headers=function()
    {
        res <- self$do_operation(options=list(select="internetMessageHeaders"))$internetMessageHeaders
        lst <- sapply(res, `[[`, "value")
        names(lst) <- sapply(res, `[[`, "name")
        lst
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
        cat("<Outlook email>\n", sep="")
        cat("  directory id:", self$properties$id, "\n")
        if(!is_empty(self$properties$from))
            cat("  from:", format_email_recipient(self$properties$from), "\n")
        else cat("  from:\n")

        if(!is_empty(self$properties$isDraft) && !self$properties$isDraft)
            cat("  sent:", format_email_date(self$properties$sentDateTime), "\n")
        else cat("  sent:\n")

        to_fmt <- sapply(self$properties$toRecipients, format_email_recipient)
        cat("  to:", paste(to_fmt, collapse=", "), "\n")

        cc_fmt <- sapply(self$properties$ccRecipients, format_email_recipient)
        if(!is_empty(cc_fmt))
            cat("  cc:", paste(cc_fmt, collapse=", "), "\n")

        bcc_fmt <- sapply(self$properties$bccRecipients, format_email_recipient)
        if(!is_empty(bcc_fmt))
            cat("  bcc:", paste(bcc_fmt, collapse=", "), "\n")

        cat("  subject:", self$properties$subject, "\n")
        cat("---\n")

        if(!is.null(self$properties$bodyPreview))
        {
            preview <- substr(self$properties$bodyPreview, 1, 255)
            if(nchar(self$properties$bodyPreview) >= 255)
                cat(preview, "...\n")
            else cat(preview, "\n")
        }
        invisible(self)
    }
),

private=list(

    make_attachment=function(object, inline, type, expiry, password, scope)
    {
        if(is.character(object) && file.exists(object) && !dir.exists(object))
        {
            # simple attachment if file is small enough, otherwise use upload session
            out <- if(is_small_attachment(file.size(object)))
            {
                list(
                    `@odata.type`="#microsoft.graph.fileAttachment",
                    isInline=inline,
                    contentBytes=openssl::base64_encode(readBin(object, "raw", file.size(object))),
                    name=basename(object),
                    contentType=mime::guess_type(object)
                )
            }
            else
            {
                if(inline)
                    stop("Inline images must be < 3MB", call.=FALSE)
                make_large_attachment(object, self)
            }
            return(out)
        }
        if(inherits(object, "ms_drive_item"))  # special treatment for OneDrive/SharePoint links
        {
            if(type == "embed")
                stop("Share link type must be one of 'view' or 'edit'", call.=FALSE)
            provider <- if(object$properties$parentReference$driveType == "personal")
                "oneDriveConsumer"
            else "oneDriveBusiness"
            permission <- paste0(scope, type)
            folder <- object$is_folder()
            object <- object$create_share_link(type, expiry, password, scope)
        }
        else
        {
            provider <- "other"
            permission <- "other"
            folder <- FALSE
        }
        if(!is.character(object) || is_empty(httr::parse_url(object)$scheme))
            stop("Attachment must be an ms_drive_item object, filename or URL", call.=FALSE)

        url <- httr::parse_url(object)
        name <- basename(url$path)
        if(name == "")
            name <- url$hostname
        list(
            `@odata.type`="#microsoft.graph.referenceAttachment",
            name=name,
            sourceUrl=object,
            isInline=inline,
            providerType=provider,
            permission=permission,
            isFolder=folder
        )
    }
))


format_email_recipient <- function(obj)
{
    name <- obj$emailAddress$name
    addr <- obj$emailAddress$address
    name_null <- is_empty(name) || nchar(name) == 0
    addr_null <- is_empty(addr) || nchar(addr) == 0

    if(addr_null) "<unknown>"
    else if(!name_null && name != addr)
        sprintf("%s <%s>", name, addr)
    else addr
}


format_email_date <- function(datestr)
{
    date <- as.POSIXct(datestr, format="%Y-%m-%dT%H:%M:%OS", tz="UTC")
    format(date, tz="", usetz=TRUE)
}


make_reply_comment <- function(comment)
{
    UseMethod("make_reply_comment")
}


make_reply_comment.default <- function(comment)
{
    as.character(comment)
}


make_reply_comment.blastula_message <- function(comment)
{
    comment$html_str
}


make_reply_comment.envelope <- function(comment)
{
    parts <- comment$parts
    inline <- which(sapply(parts, function(p) p$disposition == "inline"))
    if(length(inline) > 1)
        warning("Multiple inline sections found, only the first will be used", call.=FALSE)
    req <- if(!is_empty(inline))
        parts[[inline[1]]]$body
    else ""
}

