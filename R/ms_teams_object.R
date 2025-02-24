# Prefer header needed to work with shared channels
ms_teams_object <- R6::R6Class("ms_teams_object", inherit=ms_object,

public=list(

    do_operation=function(op="", ...)
    {
        outlook_headers <- httr::add_headers(Prefer="include-unknown-enum-members")
        super$do_operation(op, ..., outlook_headers)
    }
))
