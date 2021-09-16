tenant <- Sys.getenv("AZ_TEST_TENANT_ID")
app <- Sys.getenv("AZ_TEST_NATIVE_APP_ID")
team_name <- Sys.getenv("AZ_TEST_TEAM_NAME")
team_id <- Sys.getenv("AZ_TEST_TEAM_ID")

if(tenant == "" || app == "" || team_name == "" || team_id == "")
    skip("Teams authentication tests skipped: Microsoft Graph credentials not set")

if(!interactive())
    skip("Teams authentication tests skipped: must be in interactive session")

tok <- get_test_token(tenant, app, c("Group.ReadWrite.All", "Directory.Read.All"))
if(is.null(tok))
    skip("Teams authentication tests skipped: no access to tenant")

team <- try(call_graph_endpoint(tok, file.path("teams", team_id)), silent=TRUE)
if(inherits(team, "try-error"))
    skip("Teams authentication tests skipped: service not available")

test_that("Teams authentication works",
{
    team <- get_team(team_id=team_id, token=tok)
    expect_is(team, "ms_team")

    teamlist <- list_teams(token=tok)
    expect_is(teamlist, "list")

    teamlist2 <- list_teams(tenant=tenant, app=app)
    expect_is(teamlist2, "list")
    expect_true(all(sapply(teamlist2, inherits, "ms_team")))
    expect_true(all(mapply(
        function(s1, s2) identical(s1$properties$id, s2$properties$id,
        teamlist1,
        teamlist2
    ))))

    team2 <- get_team(team_name=team_name, token=tok)
    expect_is(team2, "ms_team")
    expect_identical(team$properties$id, team2$properties$id)

    team3 <- get_team(team_id=team_id, tenant=tenant, app=app)
    expect_is(team3, "ms_team")
    expect_identical(team$properties$id, team3$properties$id)

    team4 <- get_team(team_name=team_name, tenant=tenant, app=app)
    expect_is(team4, "ms_team")
    expect_identical(team$properties$id, team4$properties$id)
})
