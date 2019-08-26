var AVAILABILITY_BOOSTS = {
  holiday: 1,
  pto: 1,
  work_week: 1
};
var SPRINT_NAMES = [];
var SPRINT_LENGTH = 10; // days.

function getStaticDataSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Static data");
}

function getTeam() {
  var sheet = getStaticDataSheet();
  var endRow = sheet.createTextFinder("Team compositions").findNext().getRow() - 1;
  var team = {};
  for (var row = 4; row < endRow; ++row) {
    team[sheet.getRange(row, 3).getValue().trim()] = {
      name: sheet.getRange(row, 2).getValue().trim(),
      country: sheet.getRange(row, 4).getValue().trim(),
      workWeek: sheet.getRange(row, 5).getValue()
    };
  }
  addAvailabilityData(team)
  return team;
}

function getSprints() {
  if (SPRINT_NAMES.length) {
    return SPRINT_NAMES;
  }

  var sheet = getStaticDataSheet();
  var startRange = sheet.createTextFinder("Team compositions").findNext();
  var startRow = startRange.getRow() + 2;
  var startColumn = 2;
  var endColumn, value;
  var currentColumn = startColumn;
  while (!endColumn) {
    value = sheet.getRange(startRow, currentColumn).getValue().trim();
    if (value) {
      ++currentColumn;
      continue;
    }
    endColumn = currentColumn;
  }
  for (var column = startColumn; column < endColumn; ++column) {
    SPRINT_NAMES.push(sheet.getRange(startRow, column).getValue().trim());
  }
  return SPRINT_NAMES;
}

function addAvailabilityData(team) {
  var sheet = getStaticDataSheet();
  var startRow = sheet.createTextFinder("PTO Data").findNext().getRow() + 3;
  var sprints = getSprints();
  var countryData = getAvailabilityPerCountry(team, sprints);

  var assignees = Object.keys(team);
  for (var member, factors, i = 0, l = assignees.length; i < l; ++i) {
    member = team[assignees[i]];
    member.availability = (member.workWeek / 40) * AVAILABILITY_BOOSTS.work_week;

    member.sprintsAvailability = {};
    for (var daysLeft, sprint, holidays, j = 0, l2 = sprints.length; j < l2; ++j) {
      daysLeft = SPRINT_LENGTH * member.availability;
      sprint = sprints[j];
      holidays = countryData[member.country] && countryData[member.country][sprint];
      if (holidays) {
        daysLeft -= holidays * AVAILABILITY_BOOSTS.holiday;
      }
      daysLeft -= (sheet.getRange(startRow + i, j + 2).getValue() || 0) * AVAILABILITY_BOOSTS.pto;
      member.sprintsAvailability[sprint] = Math.max(daysLeft, 0) / SPRINT_LENGTH;
    }
  }

  return team;
}

function getAvailabilityPerCountry(team, sprints) {
  var sheet = getStaticDataSheet();
  
  var assignees = Object.keys(team);
  var countries = [];
  for (var member, i = 0, l = assignees.length; i < l; ++i) {
    country = team[assignees[i]].country;
    if (countries.indexOf(country) == -1) {
      countries.push(country);
    }
  }

  var startRow = sheet.createTextFinder("PTO Data").findNext().getRow() + Object.keys(team).length + 5;
  var endRow = startRow + countries.length;
  var countryData = {};
  for (var countryName, j = startRow; j < endRow; ++j) {
    countryName = sheet.getRange(j, 1).getValue().trim();
    if (countries.indexOf(countryName) == -1) {
      continue;
    }
    countryData[countryName] = {};
    for (var k = 0, l2 = sprints.length; k < l2; ++k) {
      countryData[countryName][sprints[k]] = (sheet.getRange(j, k + 2).getValue() || 0);
    }
  }
  return countryData;
}

function getAvailability(name, sprint, team) {
  var assignees = Object.keys(team);
  for (var assignee, i = 0, l = assignees.length; i < l; ++i) {
    assignee = assignees[i];
    if (team[assignee].name != name) {
      continue;
    }
    return team[assignee].sprintsAvailability[sprint] || 1;
  }
  return 1;
}

function correctCommitments(teams) {
  var membersChecked = {};
  for (var teamMembers, i = 0, l = teams.commitments.length; i < l; ++i) {
    teamMembers = Object.keys(teams.commitments[i]);
    for (var name, commitment, j = 0, l2 = teamMembers.length; j < l2; ++j) {
      name = teamMembers[j];
      commitment = teams.commitments[i][name];
      if (membersChecked[name]) {
        continue;
      } else {
        membersChecked[name] = 200 - commitment;
      }

      for (var remainder, k = 0, l3 = teams.commitments.length; k < l3; ++k) {
        if (teams.commitments[k] == teams.commitments[i] || !teams.commitments[k][name]) {
          continue;
        }
        // Adjust commitment of a team member in other teams accordingly.
        remainder = membersChecked[name] - 100;
        teams.commitments[k][name] = Math.min(teams.commitments[k][name], remainder);
        membersChecked[name] -= teams.commitments[k][name];
      }
    }
  }
  return teams;
}

function parseTeamsLine(line, allTeamNames) {
  var teams = { active: [], commitments: [] };
  var teamParts = line.split(/\]\s*,\s*\[/);
  for (var team, members, i = 0, l = teamParts.length; i < l; ++i) {
    members = teamParts[i].trim().replace(/^\[|\]$/g, "").split(",");
    for (var memberDetails, name, commitment, j = 0, l2 = members.length; j < l2; ++j) {
      memberDetails = members[j].trim().split("-");
      if (!team) {
        team = {};
      }
      team[memberDetails[0].trim()] = memberDetails.length > 1 ? parseInt(memberDetails[1], 10) : 100;
    }
    if (team) {
      var teamName = Object.keys(team).sort().join(", ");
      teams.active.push(teamName);
      if (allTeamNames.indexOf(teamName) == -1) {
        allTeamNames.push(teamName);
      }
      teams.commitments.push(team);
    }
    team = null;
  }
  return correctCommitments(teams);
}

function getTeamData() {
  var sheet = getStaticDataSheet();
  var dataRow = sheet.createTextFinder("PTO data").findNext().getRow() - 2;

  var sprints = getSprints();
  var sprintData = {};
  var allTeamNames = [];
  for (var i = 0, l = sprints.length; i < l; ++i) {
    sprintData[sprints[i]] = parseTeamsLine(sheet.getRange(dataRow, i + 2).getValue().trim(), allTeamNames);
  }

  return {
    sprints: sprintData,
    names: allTeamNames.sort().sort(function(a, b) {
      return (b.match(/,/g) || []).length - (a.match(/,/g) || []).length;
    })
  };
}

function getTeamCommitments(teams, assignees){
  for (var team, present, i = 0, l = teams.length; i < l; ++i) {
    team = teams[i];
    if (Object.keys(team).length != assignees.length) {
      continue;
    }

    present = [];
    for (var j = 0, l2 = assignees.length; j < l2; ++j) {
      if (!team[assignees[j]]) {
        break;
      }
      present.push(assignees[j]);
    }
    if (present.length == assignees.length) {
      return team;
    }
  }

  // Return something ridiculous so that the table displays weird results.
  var fakeCommitments = {};
  for (var k = 0, l3 = assignees.length; k < l3; ++k) {
    fakeCommitments[assignees[k]] = 100000;
  }
  return fakeCommitments;
}
