#' Global options
#'
#' Microsoft365R has a number of global options that affect how it interacts with the underlying Graph API.
#'
#' @section Usage:
#' ```
#' options(microsoft365r_use_itemid_in_path = TRUE)
#' options(microsoft365r_use_outlook_immutable_ids = TRUE)
#' ```
#' @section Details:
#' The `microsoft365r_use_itemid_in_path` option controls when to use item IDs in requests for OneDrive/SharePoint drive items. The default value of TRUE means to use this always; other possible values are FALSE (the default in previous versions of Microsoft365R) and "remote" (use only when dealing with items shared by another user).
#'
#' The `microsoft365r_use_outlook_immutable_ids` option controls whether to use immutable object IDs in Outlook. Immutable IDs have the advantage that they don't change when an email is moved or copied between folders, whereas traditional Outlook object IDs can change. The default is to use immutable IDs; set this option to FALSE to revert to traditional Outlook IDs.
#'
#' @name microsoft365r_options
#' @aliases microsoft365r_global
NULL

#' @import AzureGraph
NULL

utils::globalVariables(c("self", "private"))

.onLoad <- function(libname, pkgname)
{
    # set Graph API to beta, for more powerful permissions
    options(azure_graph_api_version="beta")

    # whether to use item IDs in OD/SPO paths: values are TRUE, FALSE, "remote"
    options(microsoft365r_use_itemid_in_path=TRUE)

    # whether to use immutable message IDs in Outlook
    options(microsoft365r_use_outlook_immutable_ids=TRUE)

    register_graph_class("site", ms_site,
        function(props) grepl("sharepoint", props$id, fixed=TRUE))

    register_graph_class("drive", ms_drive,
        function(props) !is_empty(props$driveType) && is_empty(props$parentReference))

    register_graph_class("driveItem", ms_drive_item,
        function(props) !is_empty(props$parentReference$driveId))

    register_graph_class("list", ms_list,
        function(props) !is_empty(props$list))

    register_graph_class("team", ms_team,
        function(props) "memberSettings" %in% names(props))

    register_graph_class("channel", ms_channel,
        function(props) "moderationSettings" %in% names(props))

    register_graph_class("chatMessage", ms_chat_message,
        function(props) "body" %in% names(props) && "messageType" %in% names(props))

    register_graph_class("plan", ms_plan,
        function(props) !is_empty(props$container) && props$container$type=='group')

    register_graph_class("plan_task", ms_plan_task,
        function(props) !is_empty(props$bucketId) && !is_empty(props$planId))

    register_graph_class("plan_bucket", ms_plan_bucket,
        function(props) !is_empty(props$planId) && !is_empty(props$orderHint) && !is_empty(props$name))

    register_graph_class("mailFolder", ms_outlook_folder,
        function(props) "unreadItemCount" %in% names(props))

    register_graph_class("message", ms_outlook_email,
        function(props) "bodyPreview" %in% names(props))

    register_graph_class("attachment", ms_outlook_attachment,
        function(props) "isInline" %in% names(props))

    register_graph_class("fileAttachment", ms_outlook_attachment,
        function(props) "isInline" %in% names(props))

    register_graph_class("referenceAttachment", ms_outlook_attachment,
        function(props) "isInline" %in% names(props))

    register_graph_class("itemAttachment", ms_outlook_attachment,
        function(props) "isInline" %in% names(props))

    register_graph_class("aadUserConversationMember", ms_team_member,
        function(props) "roles" %in% names(props))

    register_graph_class("listItem", ms_list_item,
        function(props) !is_empty(props$contentType$name))

    register_graph_class("chat", ms_chat,
        function(props) "chatType" %in% names(props))

    add_object_methods()
    add_graph_methods()
    add_user_methods()
    add_group_methods()
}

# default app ID
.microsoft365r_app_id <- "d44a05d5-c6a5-4bbb-82d2-443123722380"

# CLI for Microsoft 365 app ID
.cli_microsoft365_app_id <- "31359c7f-bd7e-475c-86db-fdb8c937548e"

# helper functions
error_message <- get("error_message", getNamespace("AzureGraph"))
get_confirmation <- get("get_confirmation", getNamespace("AzureGraph"))

# dummy mention to keep CRAN happy
# we need to ensure that vctrs is loaded so that AzureGraph will use vec_rbind
# to combine paged results into a single data frame: individual pages can have
# different structures, which will break base::rbind
vctrs::vec_rbind
