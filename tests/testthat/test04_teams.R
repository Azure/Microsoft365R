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

test_that("Teams client works",
{
    expect_error(get_team(team_name=team_name, team_id=team_id, tenant=tenant, app=app))

    team1 <- try(get_team(team_name=team_name, tenant=tenant, app=app), silent=TRUE)
    if(inherits(team1, "try-error"))
        skip("SharePoint tests skipped: service not available")

    expect_is(team1, "ms_team")
    expect_identical(team1$properties$displayName, team_name)

    team2 <- get_team(team_id=team_id, tenant=tenant, app=app)
    expect_is(team2, "ms_team")
    expect_identical(team1$properties$id, team_id)

    expect_identical(team1$properties, team2$properties)

    teams <- list_teams()
    expect_is(teams, "list")
    expect_true(all(sapply(teams, inherits, "ms_team")))
})

test_that("Teams methods work",
{
    team <- get_team(team_name, tenant=tenant, app=app)
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

    chan1 <- team$get_channel(channel_name=channel_name)
    expect_is(chan1, "ms_channel")

    chan2 <- team$get_channel(channel_id=channel_id)
    expect_is(chan2, "ms_channel")
})
