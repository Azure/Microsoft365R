# methods for different email formats: default, blastula, emayili

build_email_request <- function(body, ...)
{
    UseMethod("build_email_body")
}


build_email_body.character <- function(body, content_type, attachments, subject, to, cc, bcc, ...)
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

    utils::modifyList(req, build_email_recipients(to, cc, bcc))
}


build_email_body.blastula_message <- function(body, content_type, attachments, ...)
{
    req <- list(
        body=list(
            contentType="html",
            content=body$html_str
        )
    )
    if(!is_empty(body$attachments))
    req$attachments <- lapply(body$attachments, function(a)
    {
        assert_valid_attachment_size(a$file_path)
        list(
            `@odata.type`="#microsoft.graph.fileAttachment",
            isInline=FALSE,
            contentBytes=openssl::base64_encode(readBin(a$file_path, "raw", file.size(a$file_path))),
            name=a$filename,
            contentType=a$content_type
        )
    })
    req
}


build_email_body.envelope <- function(body, content_type, attachments, subject, to, cc, bcc, ...)
{

}


make_email_attachment <- function(object)
{
    if(!is.character(object))
        stop("Attachments must be provided as filenames or URLs", call.=FALSE)

    if(file.exists(object))  # a file
    {
        assert_valid_attachment_size(object)
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
        list(
            `@odata.type`="#microsoft.graph.referenceAttachment",
            name=basename(url$path),
            sourceUrl=object
        )
    }
    else stop("Bad attachment: '", object, "'", call.=FALSE)
}


assert_valid_attachment_size <- function(filename)
{
    if(file.size(filename) >= 3145728)
        stop("File attachments must currently be less than 3MB in size", call.=FALSE)
}
