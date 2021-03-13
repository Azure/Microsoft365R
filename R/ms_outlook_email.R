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

    set_subject=function(subject)
    {
        self$update(subject=subject)
    },

    set_recipients=function(to=NULL, cc=NULL, bcc=NULL)
    {
        if(is_empty(to) && is_empty(cc) && is_empty(bcc))
            message("Clearing all recipients")
        do.call(self$update, build_email_recipients(to, cc, bcc))
    },

    add_recipients=function(to=NULL, cc=NULL, bcc=NULL)
    {
        find_address <- function(x)
        {
            x$emailAddress$address
        }
        current_to <- sapply(self$properties$toRecipients, find_address)
        current_cc <- sapply(self$properties$ccRecipients, find_address)
        current_bcc <- sapply(self$properties$bccRecipients, find_address)

        self$set_recipients(c(current_to, to), c(current_cc, cc), c(current_bcc, bcc))
    },

    add_attachment=function(object)
    {
        res <- self$do_operation("attachments", body=make_email_attachment(object), http_verb="POST")
        ms_outlook_attachment$new(self$token, self$tenant, res,
            user_id=self$user_id, message_id=self$properties$id)
    },

    get_attachment=function(attachment_name=NULL, attachment_id=NULL)
    {
        assert_one_arg(attachment_name, attachment_id, msg="Supply exactly one of attachment name or ID")
        if(is.null(attachment_id))
        {
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

    list_attachments=function()
    {
        fields <- c("id", "name", "contentType", "size", "isInline", "lastModifiedDateTime")
        res <- self$do_operation("attachments", options=list(select=paste(fields, collapse=",")))
        private$init_list_objects(private$get_paged_list(res), default_generator=ms_outlook_attachment,
            user_id=self$user_id, message_id=self$properties$id)
    },

    remove_attachment=function(attachment_name, attachment_id, confirm=TRUE)
    {
        self$get_attachment(attachment_name, attachment_id)$delete(confirm=confirm)
    },

    download_attachment=function(attachment_name, attachment_id, dest, overwrite=FALSE)
    {
        self$get_attachment(attachment_name, attachment_id)$download(dest, overwrite=FALSE)
    },

    send=function()
    {
        self$do_operation("send", http_verb="POST")
        self$sync_fields()
    },

    get_message_headers=function()
    {
        res <- self$do_operation(options=list(select="internetMessageHeaders"))$internetMessageHeaders
        lst <- sapply(res, `[[`, "value")
        names(lst) <- sapply(res, `[[`, "name")
        lst
    },

    copy=function(folder_name, folder_id)
    {
        assert_one_arg(folder_name, folder_id, msg="Supply exactly one of destination folder name or ID")
        if(is.null(folder_id))
            folder_id <- private$find_folder_id(folder_name)

        res <- self$do_operation("copy", body=list(destinationId=folder_id), http_verb="POST")
        ms_outlook_email$new(self$token, self$tenant, res, user_id=self$user_id)
    },

    move=function(folder_name, folder_id)
    {
        assert_one_arg(folder_name, folder_id, msg="Supply exactly one of destination folder name or ID")
        if(is.null(folder_id))
            folder_id <- private$find_folder_id(folder_name)

        res <- self$do_operation("move", body=list(destinationId=folder_id), http_verb="POST")
        ms_outlook_email$new(self$token, self$tenant, res, user_id=self$user_id)
    },

    reply_all=function(comment="")
    {
        op <- "createReplyAll"
        body <- list(comment=comment)
        reply <- ms_outlook_email$new(self$token, self$tenant,
            self$do_operation("op", body=body, http_verb="POST"), user_id=self$user_id)
        if(send_now)
            reply$send()
        reply
    },

    reply=function(comment="", send_now=FALSE)
    {
        op <- "createReply"
        body <- list(comment=comment)
        reply <- ms_outlook_email$new(self$token, self$tenant,
            self$do_operation("op", body=body, http_verb="POST"), user_id=self$user_id)
        if(send_now)
            reply$send()
        reply
    },

    forward=function(comment="", to=NULL, cc=NULL, bcc=NULL, send_now=FALSE)
    {
        op <- "createforward"
        body <- c(
            list(comment=comment),
            build_email_recipients(to, cc, bcc)
        )
        reply <- ms_outlook_email$new(self$token, self$tenant,
            self$do_operation("op", body=list(comment=comment), http_verb="POST"), user_id=self$user_id)
        if(send_now)
            reply$send()
        reply
    },

    print=function(...)
    {
        cat("<Outlook email>\n", sep="")
        cat("  directory id:", self$properties$id, "\n")
        if(!is_empty(self$properties$from))
            cat("  from:", format_email_recipient(self$properties$from), "\n")
        else cat("  from:\n")

        if(!self$properties$isDraft)
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

        cat(self$properties$bodyPreview)
        if(nchar(self$properties$bodyPreview) >= 255)
            cat(" ...\n")
        else cat("\n")
        invisible(self)
    }
),

private=list(

    find_folder_id=function(folder)
    {
        if(is.character(folder))
        {
            fnames <- strsplit(folder, "/", fixed=TRUE)[[1]]
            # create dummy outlook client to get top-level folder
            outlook <- ms_outlook$new(self$token, self$tenant, list(id=self$user_id))
            folder <- outlook$get_folder(fnames[1])
            for(f in fnames[-1])
                folder <- folder$get_folder(f)
        }
        if(!inherits(folder, "ms_outlook_folder"))
            stop("Invalid folder object", call.=FALSE)
        folder$properties$id
    }
))


format_email_recipient <- function(obj)
{
    name <- obj$emailAddress$name
    addr <- obj$emailAddress$address
    name_null <- is_empty(name) || nchar(name) == 0
    addr_null <- is_empty(addr) || nchar(addr) == 0

    if(name_null && addr_null)
        "<unknown>"
    else if(name_null)
        addr
    else name
}


format_email_date <- function(datestr)
{
    date <- as.POSIXct(datestr, format="%Y-%m-%dT%H:%M:%OS", tz="UTC")
    format(date, tz="", usetz=TRUE)
}
