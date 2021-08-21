build_chatmessage_body <- function(channel, body, content_type, attachments, inline, mentions)
{
    get_upload_location <- function(item)
    {
        path <- item$get_parent_folder()$properties$webUrl
        name <- item$properties$name
        file.path(path, name)
    }

    call_body <- list(body=list(content=paste(body, collapse="\n"), contentType=content_type))
    if(!is_empty(attachments))
    {
        call_body$attachments <- lapply(attachments, function(f)
        {
            att <- channel$upload_file(f, dest=basename(f))
            et <- att$properties$eTag
            list(
                id=regmatches(et, regexpr("[A-Za-z0-9\\-]{10,}", et)),
                name=att$properties$name,
                contentUrl=get_upload_location(att),
                contentType="reference"
            )
        })
        att_tags <- lapply(call_body$attachments,
            function(att) paste0('<attachment id="', att$id, '"></attachment>'))
        call_body$body$content <- paste(call_body$body$content, paste(att_tags, collapse=""))
    }
    if(!is_empty(inline))
    {
        if(call_body$body$contentType != "html")
            stop("Content type must be 'html' to include inline content", .call=FALSE)

        call_body$hostedContents <- lapply(seq_along(inline), function(i)
        {
            f <- inline[i]
            cont <- openssl::base64_encode(readBin(f, "raw", file.size(f)))
            list(
                `@microsoft.graph.temporaryId`=as.character(i),
                contentBytes=cont,
                contentType=mime::guess_type(f)
            )
        })
        inline_tags <- lapply(seq_along(inline), function(i)
        {
            sprintf('<div><span><img src="../hostedContents/%d/$value" style="vertical-align:bottom"></span>\n</div>',
                    i)
        })
        call_body$body$content <- paste(call_body$body$content, paste(inline_tags, collapse=""))
    }
    if(!is_empty(mentions))
    {
        if(call_body$body$contentType != "html")
            stop("Content type must be 'html' to include mentions", .call=FALSE)
        if(inherits(mentions, c("ms_team_member", "az_user", "ms_team", "ms_channel")))
            mentions <- list(mentions)

        call_body$mentions <- lapply(seq_along(mentions), function(i)
        {
            obj <- mentions[[i]]
            if(!inherits(obj, c("ms_team_member", "az_user", "ms_team", "ms_channel")))
                stop("Must supply an object representing a team member, user, team or channel", call.=FALSE)
            make_mention(obj, i)
        })
        mention_tags <- lapply(call_body$mentions,
            function(m) sprintf('<at id="%d">%s</at>', m$id, m$mentionText))
        call_body$body$content <- paste(call_body$body$content, paste(mention_tags, collapse=" "))
    }
    xx <<- call_body
    call_body
}


make_mention <- function(object, i)
{
    UseMethod("make_mention")
}


make_mention.az_user <- function(object, i)
{
    name <- if(!is.null(object$properties$displayName))
        object$properties$displayName
    else if(!is.null(object$properties$userPrincipalName))
        object$properties$userPrincipalName
    else stop("Could not find user display name", call.=FALSE)
    list(
        id=i,
        mentionText=name,
        mentioned=list(
            user=list(
                id=object$properties$id,
                displayName=object$properties$displayName,
                userIdentityType="aadUser"
            )
        )
    )
}


make_mention.ms_team <- function(object, i)
{
    list(
        id=i,
        mentionText=object$properties$displayName,
        mentioned=list(
            conversation=list(
                id=object$properties$id,
                displayName=object$properties$displayName,
                conversationIdentityType="team"
            )
        )
    )
}


make_mention.ms_channel <- function(object, i)
{
    list(
        id=i,
        mentionText=object$properties$displayName,
        mentioned=list(
            conversation=list(
                id=object$properties$id,
                displayName=object$properties$displayName,
                conversationIdentityType="channel"
            )
        )
    )
}


make_mention.ms_team_member <- function(object, i)
{
    make_mention(object$get_aaduser(), i)
}
