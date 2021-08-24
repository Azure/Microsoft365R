tenant <- Sys.getenv("AZ_TEST_TENANT_ID")
app <- Sys.getenv("AZ_TEST_NATIVE_APP_ID")
team_name <- Sys.getenv("AZ_TEST_TEAM_NAME")
team_id <- Sys.getenv("AZ_TEST_TEAM_ID")
mention_name <- Sys.getenv("AZ_TEST_CHANNEL_MENTION_NAME")

if(tenant == "" || app == "" || team_name == "" || team_id == "" || mention_name == "")
    skip("Channel tests skipped: Microsoft Graph credentials not set")

if(Sys.getenv("AZ_TEST_CHANNEL_FLAG") == "")
    skip("Channel tests skipped: flag not set")

if(!interactive())
    skip("Channel tests skipped: must be in interactive session")

tok <- try(AzureAuth::get_azure_token(
    c("https://graph.microsoft.com/.default",
      "openid",
      "offline_access"),
    tenant=tenant, app=app, version=2),
    silent=TRUE)
if(inherits(tok, "try-error"))
    skip("Channel tests skipped: no access to tenant")

team <- try(call_graph_endpoint(tok, file.path("teams", team_id)), silent=TRUE)
if(inherits(team, "try-error"))
    skip("Channel tests skipped: service not available")

channel_name <- sprintf("Test channel %s", make_name(10))

test_that("Channel methods work",
{
    team <- get_team(team_id=team_id, tenant=tenant, app=app)
    expect_is(team, "ms_team")

    expect_error(team$get_channel(channel_name=channel_name))

    chan <- team$create_channel(channel_name, description="Temporary testing channel", membership="private")
    expect_is(chan, "ms_channel")
    expect_false(inherits(chan$properties, "xml_document"))

    Sys.sleep(10)
    folder <- chan$get_folder()
    expect_is(folder, "ms_drive_item")

    lst <- chan$list_messages()
    expect_is(lst, "list")
    expect_identical(length(lst), 0L)

    msg_body <- sprintf("Test message: %s", make_name(5))
    msg <- chan$send_message(msg_body)
    expect_is(msg, "ms_chat_message")

    msg2_body <- sprintf("<div>Test message: %s</div", make_name(5))
    msg2 <- chan$send_message(msg2_body, content_type="html")
    expect_is(msg2, "ms_chat_message")

    msg3_body <- sprintf("Test message with attachment: %s", make_name(5))
    f0 <- write_file(fname=tempfile(tmpdir=tempdir(), fileext=".xlsx"))
    msg3 <- chan$send_message(msg3_body, attachments=f0)
    expect_is(msg3, "ms_chat_message")
    expect_true(nzchar(msg3$properties$attachments[[1]]$contentUrl))

    msg4_body <- sprintf("Test message with inline image: %s", make_name(5))
    expect_error(chan$send_message(msg4_body, inline="../resources/logo_small.jpg"))
    msg4 <- chan$send_message(msg4_body, content_type="html", inline="../resources/logo_small.jpg")
    expect_is(msg4, "ms_chat_message")

    repl_body <- sprintf("Test reply: %s", make_name(5))
    repl <- msg$send_reply(repl_body)
    expect_is(repl, "ms_chat_message")

    expect_error(repl$send_reply("Reply to reply"))

    # expect_silent(msg$delete(confirm=FALSE))
    # expect_silent(chan$delete_message(msg2$properties$id, confirm=FALSE))
    # expect_silent(chan$delete_message(msg3$properties$id, confirm=FALSE))

    f1 <- write_file()
    it <- chan$upload_file(f1)
    expect_is(it, "ms_drive_item")

    flist <- chan$list_files(info="name")
    expect_true(basename(f0) %in% flist)
    expect_true(basename(f1) %in% flist)

    f_dl <- tempfile()
    expect_silent(chan$download_file(basename(f1), f_dl))
    expect_true(files_identical(f1, f_dl))

    # members, message mentions
    mlst <- chan$list_members()
    expect_is(mlst, "list")
    expect_true(all(sapply(mlst, inherits, "ms_team_member")))
    expect_true(all(sapply(mlst, function(obj) obj$type == "channel member")))

    usr <- chan$get_member(mention_name)
    usrname <- usr$properties$displayName
    usremail <- usr$properties$email
    usrid <- usr$properties$id
    expect_false(is.null(usrname))
    expect_false(is.null(usremail))
    expect_false(is.null(usrid))

    mmsg <- chan$send_message("Mention test", content_type="html", mentions=usr)
    expect_is(mmsg, "ms_chat_message")
    expect_false(is.null(mmsg$properties$mentions))

    mmsg2 <- chan$send_message("Mention test 2", content_type="html", mentions=list(chan, team))
    expect_is(mmsg2, "ms_chat_message")
    expect_false(is.null(mmsg2$properties$mentions))

    lapply(flist, function(f) folder$get_item(f)$delete(confirm=FALSE))
    expect_silent(chan$delete(confirm=FALSE))
})
