# allow use of immutable object IDs when talking to Outlook
ms_outlook_object <- R6::R6Class("ms_outlook_object", inherit=ms_object,

public=list(

    do_operation=function(op="", ...)
    {
        if(getOption("microsoft365r_use_outlook_immutable_ids"))
        {
            outlook_headers <- httr::add_headers(Prefer='IdType="ImmutableId"')
            super$do_operation(op, ..., outlook_headers)
        }
        else super$do_operation(op, ...)
    }
))
