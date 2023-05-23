add_external_attachments <- function(object, email)
{
    UseMethod("add_external_attachments")
}


add_external_attachments.blastula_message <- function(object, email)
{
    for(a in object$attachments)
        email$add_attachment(a$file_path)

    for(i in seq_along(object$images))
    {
        if(!is_small_attachment(nchar(object$images[[i]])/0.74))  # allow for base64 bloat
        {
            warning("Inline images must be < 3MB; will be skipped", call.=FALSE)
            next
        }
        body <- list(
            `@odata.type`="#microsoft.graph.fileAttachment",
            contentBytes=object$images[[i]],
            contentId=names(object$images)[[i]],
            name=names(object$images)[[i]],
            contentType=attr(object$images[[i]], "content_type"),
            isInline=TRUE
        )
        email$do_operation("attachments", body=body, http_verb="POST")
    }
}


add_external_attachments.envelope <- function(object, email)
{
    require_emayili_0.6()
    parts <- object$parts

    # parts is either a single body object (itself a named list), or a list of body objects
    if(!is.null(names(parts)))
        parts <- list(parts)

    atts <-  which(sapply(parts, inherits, "attachment"))
    for(a in parts[atts])
    {
        if(!is_small_attachment(length(a$content)))
        {
            warning("File attachments from emayili > 3MB not currently supported; will be skipped", call.=FALSE)
            next
        }

        name_a <- unclass(gsub('"', "", sub("^.+filename=", "", a$disposition)))
        att <- list(
            `@odata.type`="#microsoft.graph.fileAttachment",
            isInline=FALSE,
            contentBytes=openssl::base64_encode(a$content),
            contentId=name_a,
            name=name_a,
            contentType=unclass(sub(";.+$", "", a$type))
        )
        email$do_operation("attachments", body=att, http_verb="POST")
    }
}


add_external_attachments.default <- function(object, email)
{
    # do nothing if message object is not a recognised class (from blastula or emayili)
    NULL
}


is_small_attachment <- function(filesize, threshold=3e6)
{
    filesize <= threshold
}


make_large_attachment <- function(object, email)
{
    if(missing(email) || !inherits(email, "ms_outlook_email"))
        stop("Must supply email object", call.=FALSE)

    size <- file.size(object)
    body <- list(attachmentItem=list(
        attachmentType="file",
        name=basename(object),
        size=size
    ))
    upload_dest <- email$do_operation("attachments/createUploadSession", body=body, http_verb="POST")$uploadUrl

    con <- file(object, open="rb")
    on.exit(close(con))
    next_blockstart <- 0
    next_blockend <- size - 1
    blocksize <- 3145728
    repeat
    {
        next_blocksize <- min(next_blockend - next_blockstart + 1, blocksize)
        seek(con, next_blockstart)
        body <- readBin(con, "raw", next_blocksize)
        thisblock <- length(body)
        if(thisblock == 0)
            break

        headers <- httr::add_headers(
            `Content-Length`=thisblock,
            `Content-Range`=sprintf("bytes %.0f-%.0f/%.0f",
                next_blockstart, next_blockstart + thisblock - 1, size)
        )
        res <- httr::PUT(upload_dest, headers, body=body)
        httr::stop_for_status(res)

        next_block <- parse_upload_range(httr::content(res), blocksize)
        if(is.null(next_block))
            break
        next_blockstart <- next_block[1]
        next_blockend <- next_block[2]
    }
    invisible(NULL)
}
