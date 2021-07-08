ms_chat <- R6::R6Class("ms_chat", inherit=ms_object,

public=list(

    initialize=function(token, tenant=NULL, properties=NULL)
    {
        self$type <- "chat"
        private$api_type <- "chats"
        super$initialize(token, tenant, properties)
    },

    send_message=function(body, content_type=c("text", "html"), attachments=NULL, inline=NULL, mentions=NULL)
    {
        content_type <- match.arg(content_type)
        call_body <- build_chatmessage_body(self, body, content_type, attachments, inline, mentions)
        res <- self$do_operation("messages", body=call_body, http_verb="POST")
        ms_chat_message$new(self$token, self$tenant, res)
    },

    list_messages=function(filter=NULL, n=50)
    {
        make_basic_list(self, "messages", filter, n)
    },

    get_message=function(message_id)
    {
        op <- file.path("messages", message_id)
        ms_chat_message$new(self$token, self$tenant, self$do_operation(op))
    },

    delete_message=function(message_id, confirm=TRUE)
    {
        self$get_message(message_id)$delete(confirm=confirm)
    },

    list_members=function(filter=NULL, n=Inf)
    {
        make_basic_list(self, "members", filter, n, parent_id=self$properties$id, parent_type="chat")
    },

    get_member=function(name=NULL, email=NULL, id=NULL)
    {
        assert_one_arg(name, email, id, msg="Supply exactly one of member name, email address, or ID")
        if(!is.null(id))
        {
            res <- self$do_operation(file.path("members", id))
            ms_team_member$new(self$token, self$tenant, res,
                parent_id=self$properties$id, parent_type="chat")
        }
        else
        {
            filter <- if(!is.null(name))
                sprintf("displayName eq '%s'", name)
            else sprintf("microsoft.graph.aadUserConversationMember/email eq '%s'", email)
            res <- self$list_members(filter=filter)
            if(length(res) != 1)
                stop("Invalid name or email address", call.=FALSE)
            res[[1]]
        }
    },

    print=function(...)
    {
        type <- switch(self$properties$chatType,
            "oneOnOne"="One on one",
            "group"="Group",
            "meeting"="Meeting",
            "Other"
        )
        cat("<", type, " chat '", self$properties$displayName, "'>\n", sep="")
        cat("  directory id:", self$properties$id, "\n")
        cat("  web link:", self$properties$webUrl, "\n")
        cat("---\n")
        cat(format_public_methods(self))
        invisible(self)
    }
))
