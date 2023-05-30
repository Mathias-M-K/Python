from nba_nfl_stuff.Assignments import ShowOnlyTeamNames, CleanTeamNames

documents = ["nba", "nfl"]

for document in documents:
    CleanTeamNames.do_assignment(document)

for document in documents:
    ShowOnlyTeamNames.do_assignment(document)


