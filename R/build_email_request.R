# methods for different email formats: default, blastula, emayili

build_email_request <- function(body, ...)
{
    UseMethod("build_email_request")
}


build_email_request.character <- function(body, content_type,
    attachments=NULL, subject=NULL, to=NA, cc=NA, bcc=NA, reply_to=NA, ...)
{
    req <- list(
        body=list(
            contentType=content_type,
            content=paste(body, collapse="\n")
        )
    )

    if(!is_empty(attachments))
        req$attachments <- lapply(attachments, make_email_attachment)

    if(!is_empty(subject))
        req$subject <- subject

    utils::modifyList(req, build_email_recipients(to, cc, bcc, reply_to))
}


build_email_request.blastula_message <- function(body, content_type,
    attachments=NULL, subject=NULL, to=NA, cc=NA, bcc=NA, reply_to=NA, ...)
{
    req <- list(
        body=list(
            contentType="html",  # blastula emails are always HTML
            content=body$html_str
        )
    )

    if(!is_empty(body$attachments))
        req$attachments <- lapply(body$attachments, function(a)
        {
            assert_valid_attachment_size(file.size(a$file_path))
            list(
                `@odata.type`="#microsoft.graph.fileAttachment",
                isInline=FALSE,
                contentBytes=openssl::base64_encode(readBin(a$file_path, "raw", file.size(a$file_path))),
                name=a$filename,
                contentType=a$content_type
            )
        })

    if(!is_empty(subject))
        req$subject <- subject

    utils::modifyList(req, build_email_recipients(to, cc, bcc, reply_to))
}


build_email_request.envelope <- function(body, ...)
{
    parts <- body$parts

    inline <- which(sapply(parts, function(p) p$header$content_disposition == "inline"))
    if(length(inline) > 1)
        warning("Multiple inline sections found, only the first will be used", call.=FALSE)
    req <- if(!is_empty(inline))
    {
        inline <- parts[[inline[1]]]
        list(
            body=list(
                contentType=if(inline$header$content_type == "text/html") "html" else "text",
                content=inline$body
            )
        )
    }
    else list(body=list(contentType="text", content=""))

    atts <-  which(sapply(parts, function(p) p$header$content_disposition == "attachment"))
    if(!is_empty(atts))
        req$attachments <- lapply(parts[atts], function(a)
        {
            assert_valid_attachment_size(nchar(a$body)/0.74)  # allow for base64 bloat
            list(
                `@odata.type`="#microsoft.graph.fileAttachment",
                isInline=FALSE,
                contentBytes=a$body,
                name=a$header$filename,
                contentType=a$header$content_type
            )
        })

    if(!is_empty(body$header$Subject))
        req$subject <- body$header$Subject

    utils::modifyList(req,
        build_email_recipients(body$header$To, body$header$Cc, body$header$Bcc, body$header$Reply_To))
}


make_email_attachment <- function(object)
{
    if(!is.character(object))
        stop("Attachments must be provided as filenames or URLs", call.=FALSE)

    if(file.exists(object))  # a file
    {
        assert_valid_attachment_size(file.size(object))
        list(
            `@odata.type`="#microsoft.graph.fileAttachment",
            isInline=FALSE,
            contentBytes=openssl::base64_encode(readBin(object, "raw", file.size(object))),
            name=basename(object),
            contentType=mime::guess_type(object)
        )
    }
    else if(!is_empty(httr::parse_url(object)$scheme))  # a URL
    {
        url <- httr::parse_url(object)
        name <- basename(url$path)
        if(name == "")
            name <- url$hostname
        list(
            `@odata.type`="#microsoft.graph.referenceAttachment",
            name=name,
            sourceUrl=object,
            providerType="other",
            permission="view",
            isFolder=FALSE
        )
    }
    else stop("Bad attachment: '", object, "'", call.=FALSE)
}


assert_valid_attachment_size <- function(filesize)
{
    if(filesize >= 3145728)
        stop("File attachments must currently be less than 3MB in size", call.=FALSE)
}


build_email_recipients <- function(to, cc, bcc, reply_to)
{
    make_recipients <- function(addr_list)
    {
        # NA means don't update current value
        if(!is_empty(addr_list) && is.na(addr_list))
            return(NA)

        # handle case of a single az_user object
        if(is.object(addr_list))
            addr_list <- list(addr_list)

        lapply(addr_list, function(x)
        {
            if(inherits(x, "az_user"))
            {
                props <- x$properties
                x <- if(!is.null(props$mail))
                    props$mail
                else props$userPrincipalName
                if(is_empty(x) || nchar(x) == 0)
                    stop("Unable to find email address", call.=FALSE)
                name <- props$displayName
            }
            else name <- x <- as.character(x)
            if(!grepl(".+@.+", x))  # basic check for a valid address
                stop("Invalid email address '", x, "'", call.=FALSE)
            list(emailAddress=list(name=name, address=x))
        })
    }

    out <- list(
        toRecipients=make_recipients(to),
        ccRecipients=make_recipients(cc),
        bccRecipients=make_recipients(bcc),
        replyTo=make_recipients(reply_to)
    )
    out[sapply(out, function(x) is_empty(x) || !is.na(x))]
}
