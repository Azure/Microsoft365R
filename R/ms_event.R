ms_event <- R6::R6Class("ms_event", inherit=ms_object,

public=list(

    initialize=function(token, tenant=NULL, properties=NULL)
    {
        self$type <- "Outlook event"
        private$api_type <- "events"
        super$initialize(token, tenant, properties)
    }

))
