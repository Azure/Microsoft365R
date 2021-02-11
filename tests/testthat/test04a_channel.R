tenant <- Sys.getenv("AZ_TEST_TENANT_ID")
app <- Sys.getenv("AZ_TEST_NATIVE_APP_ID")
team_name <- Sys.getenv("AZ_TEST_TEAM_NAME")
team_id <- Sys.getenv("AZ_TEST_TEAM_ID")
channel_name <- Sys.getenv("AZ_TEST_CHANNEL_NAME")
channel_id <- Sys.getenv("AZ_TEST_CHANNEL_ID")

if(tenant == "" || app == "" || team_name == "" || team_id == "" || channel_name == "" || channel_id == "")
    skip("Channel tests skipped: Microsoft Graph credentials not set")

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

test_that("Channel methods work",
{
    team <- get_team(team_id=team_id, tenant=tenant, app=app)
    expect_is(team, "ms_team")

    chan <- team$get_channel(channel_name=channel_name)
    expect_is(chan, "ms_channel")

    lst <- chan$list_messages()
    expect_is(lst, "list")
    expect_true(all(sapply(lst, inherits, "ms_chat_message")))

    msg_body <- sprintf("Test message: %s", make_name(5))
    msg <- chan$send_message(msg_body)
    expect_is(msg, "ms_chat_message")

    msg2_body <- sprintf("<div>Test message: %s</div", make_name(5))
    msg2 <- chan$send_message(msg2_body, content_type="html")
    expect_is(msg2, "ms_chat_message")

    msg3_body <- sprintf("Test message with attachment: %s", make_name(5))
    f0 <- write_file()
    msg3 <- chan$send_message(msg3_body, attachments=f0)
    expect_is(msg3, "ms_chat_message")
    expect_true(!nzchar(msg3$properties$attachments$contentUrl))

    repl_body <- sprintf("Test reply: %s", make_name(5))
    repl <- msg$send_reply(repl_body)
    expect_is(repl, "ms_chat_message")

    expect_error(repl$send_reply("Reply to reply"))

    expect_silent(msg$delete(confirm=FALSE))
    expect_silent(chan$delete_message(msg2$properties$id, confirm=FALSE))
    expect_silent(chan$delete_message(msg3$properties$id, confirm=FALSE))

    f1 <- write_file()
    it <- chan$upload_file(f1)
    expect_is(it, "ms_drive_item")

    flist <- chan$list_files(info="name")
    expect_true(basename(f0) %in% flist)
    expect_true(basename(f1) %in% flist)

    f_dl <- tempfile()
    expect_silent(chan$download_file(basename(f1), f_dl))
    expect_true(files_identical(f1, f_dl))

    drv <- team$get_drive()
    itempath0 <- file.path(channel_name, basename(f0))
    itempath1 <- file.path(channel_name, basename(f1))
    expect_silent(drv$delete_item(itempath0, confirm=FALSE))
    expect_silent(drv$delete_item(itempath1, confirm=FALSE))
})
