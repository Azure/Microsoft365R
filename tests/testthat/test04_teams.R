tenant <- Sys.getenv("AZ_TEST_TENANT_ID")
app <- Sys.getenv("AZ_TEST_NATIVE_APP_ID")
team_name <- Sys.getenv("AZ_TEST_TEAM_NAME")
team_id <- Sys.getenv("AZ_TEST_TEAM_ID")
channel_name <- Sys.getenv("AZ_TEST_CHANNEL_NAME")
channel_id <- Sys.getenv("AZ_TEST_CHANNEL_ID")

if(tenant == "" || app == "" || team_name == "" || team_id == "" || channel_name == "" || channel_id == "")
    skip("Teams tests skipped: Microsoft Graph credentials not set")

if(!interactive())
    skip("Teams tests skipped: must be in interactive session")

tok <- try(AzureAuth::get_azure_token(
    c("https://graph.microsoft.com/.default",
      "openid",
      "offline_access"),
    tenant=tenant, app=app, version=2),
    silent=TRUE)
if(inherits(tok, "try-error"))
    skip("Teams tests skipped: no access to tenant")

team <- try(call_graph_endpoint(tok, file.path("teams", team_id)), silent=TRUE)
if(inherits(team, "try-error"))
    skip("Teams tests skipped: service not available")

test_that("Teams client works",
{
    expect_error(get_team(team_name=team_name, team_id=team_id, tenant=tenant, app=app))

    team1 <- get_team(team_name=team_name, tenant=tenant, app=app)
    expect_is(team1, "ms_team")
    expect_identical(team1$properties$displayName, team_name)

    team2 <- get_team(team_id=team_id, tenant=tenant, app=app)
    expect_is(team2, "ms_team")
    expect_identical(team1$properties$id, team_id)

    teams <- list_teams()
    expect_is(teams, "list")
    expect_true(all(sapply(teams, inherits, "ms_team")))
})

test_that("Teams methods work",
{
    team <- get_team(team_id=team_id, tenant=tenant, app=app)
    expect_is(team, "ms_team")

    # drive -- functionality tested in test02
    drives <- team$list_drives()
    expect_is(drives, "list")
    expect_true(all(sapply(drives, inherits, "ms_drive")))

    drv <- team$get_drive()
    expect_is(drv, "ms_drive")

    grp <- team$get_group()
    expect_is(grp, "az_group")

    site <- team$get_sharepoint_site()
    expect_is(site, "ms_site")

    # channels
    chans <- team$list_channels()
    expect_is(chans, "list")
    expect_true(all(sapply(chans, inherits, "ms_channel")))

    expect_error(team$get_channel(channel_name, channel_id))

    chan0 <- team$get_channel()
    expect_is(chan0, "ms_channel")
    f0 <- chan0$get_folder()
    expect_is(f0, "ms_drive_item")
    expect_is(f0$list_files(), "data.frame")

    chan1 <- team$get_channel(channel_name=channel_name)
    expect_is(chan1, "ms_channel")
    f1 <- chan1$get_folder()
    expect_is(f1, "ms_drive_item")
    expect_is(f1$list_files(), "data.frame")

    src <- write_file()
    it <- chan1$upload_file(src)
    expect_is(it, "ms_drive_item")
    expect_silent(it$delete(confirm=FALSE))

    chan2 <- team$get_channel(channel_id=channel_id)
    expect_is(chan2, "ms_channel")
})

test_that("Team member methods work",
{
    team <- get_team(team_id=team_id, tenant=tenant, app=app)

    mlst <- team$list_members()
    expect_is(mlst, "list")
    expect_true(all(sapply(mlst, inherits, "ms_team_member")))
    expect_true(all(sapply(mlst, function(obj) obj$type == "team member")))

    usr <- mlst[[1]]
    usrname <- usr$properties$displayName
    usremail <- usr$properties$email
    usrid <- usr$properties$id
    expect_false(is.null(usrname))
    expect_false(is.null(usremail))
    expect_false(is.null(usrid))

    usr1 <- team$get_member(usrname)
    expect_is(usr1, "ms_team_member")
    expect_identical(usr$properties$id, usr1$properties$id)

    usr2 <- team$get_member(email=usremail)
    expect_is(usr2, "ms_team_member")
    expect_identical(usr$properties$id, usr2$properties$id)

    usr3 <- team$get_member(id=usrid)
    expect_is(usr3, "ms_team_member")
    expect_identical(usr$properties$id, usr3$properties$id)

    aaduser <- usr$get_aaduser()
    expect_is(aaduser, "az_user")

    aaduser1 <- usr1$get_aaduser()
    expect_is(aaduser1, "az_user")
})
