ms_calendar <- R6::R6Class("ms_calendar", inherit=ms_object,

public=list(

    initialize=function(token, tenant=NULL, properties=NULL)
    {
        self$type <- "Outlook calendar"
        private$api_type <- "calendar"
        super$initialize(token, tenant, properties)
    },

    list_events=function(filter=NULL, n=Inf)
    {
        private$make_basic_list("events", filter, n)
    },

    get_event=function()
    {},

    create_event=function(...)
    {},

    update_event=function(...)
    {},

    delete_event=function(id, confirm=TRUE)
    {},

    print=function(...)
    {
        owner <- self$properties$owner
        if(!is_empty(owner))
        {
            name <- owner$name
            email <- owner$email
        }
        else name <- email <- NA_character_
        cat("<Outlook calendar for '", name, "'>\n", sep="")
        cat("  email address:", email, "\n")
        cat("---\n")
    }
))

