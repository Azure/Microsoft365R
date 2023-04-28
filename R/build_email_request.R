# methods for different email formats: default, blastula, emayili

build_email_request <- function(body, ...)
{
    UseMethod("build_email_request")
}


build_email_request.character <- function(body, content_type,
    subject=NULL, to=NA, cc=NA, bcc=NA, reply_to=NA, token=NULL, user_id=NULL, ...)
{
    req <- list(
        body=list(
            contentType=content_type,
            content=paste(body, collapse="\n")
        )
    )
    if(!is_empty(subject))
        req$subject <- subject

    utils::modifyList(req, build_email_recipients(to, cc, bcc, reply_to))
}


build_email_request.blastula_message <- function(body, content_type,
    subject=NULL, to=NA, cc=NA, bcc=NA, reply_to=NA, token=NULL, user_id=NULL, ...)
{
    req <- list(
        body=list(
            contentType="html",  # blastula emails are always HTML
            content=body$html_str
        )
    )
    if(!is_empty(subject))
        req$subject <- subject

    utils::modifyList(req, build_email_recipients(to, cc, bcc, reply_to))
}


build_email_request.envelope <- function(body, token=NULL, user_id=NULL, ...)
{
    require_emayili_0.6()
    parts <- body$parts

    # parts is either a single body object (itself a named list), or a list of body objects
    if(!is.null(names(parts)))
        parts <- list(parts)

    inline <- which(sapply(parts, function(p) p$disposition == "inline"))
    if(length(inline) > 1)
        warning("Multiple inline sections found, only the first will be used", call.=FALSE)
    req <- if(!is_empty(inline))
    {
        inline <- parts[[inline[1]]]
        list(
            body=list(
                contentType=if(inherits(inline, "text_html")) "html" else "text",
                content=inline$content
            )
        )
    }
    else list(body=list(contentType="text", content=""))

    if(!is_empty(emayili::subject(body)))
        req$subject <- as.character(emayili::subject(body))

    recipients <- build_email_recipients(
        as.character(emayili::to(body)),
        as.character(emayili::cc(body)),
        as.character(emayili::bcc(body)),
        as.character(emayili::reply(body))
    )

    utils::modifyList(req, recipients)
}


build_email_recipients <- function(to, cc, bcc, reply_to)
{
    make_recipients <- function(addr_list)
    {
        # NA means don't update current value
        if(!is_empty(addr_list) && any(is.na(addr_list)))
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
            if(!all(grepl(".+@.+", x)))  # basic check for a valid address
                stop("Invalid email address supplied", call.=FALSE)
            list(emailAddress=list(name=name, address=x))
        })
    }

    out <- list(
        toRecipients=make_recipients(to),
        ccRecipients=make_recipients(cc),
        bccRecipients=make_recipients(bcc),
        replyTo=make_recipients(reply_to)
    )
    out[sapply(out, function(x) is_empty(x) || all(!is.na(x)))]
}


require_emayili_0.6 <- function()
{
    if(utils::packageVersion("emayili") < package_version("0.6"))
        stop("Need emayili version 0.6 or later", call.=FALSE)
}
