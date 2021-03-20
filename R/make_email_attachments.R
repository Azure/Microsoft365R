add_external_attachments <- function(object, email)
{
    UseMethod("add_external_attachments")
}


add_external_attachments.blastula_message <- function(object, email)
{

}


add_external_attachments.envelope <- function(object, email)
{

}


add_external_attachments.default <- function(object, email)
{
    NULL
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
