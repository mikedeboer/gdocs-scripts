var AVAILABILITY_BOOSTS = {
  holiday: 1,
  pto: 1,
  work_week: 1
};
var SPRINT_NAMES = [];
var SPRINT_LENGTH = 10; // days.
var CELL_VALUE_CACHE = [];

function getStaticDataSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Static data");
}

function getStaticCellValue(sheet, row, column, emptyCellValue) {
  if (typeof emptyCellValue == "undefined") {
    emptyCellValue = "";
  }
  if (CELL_VALUE_CACHE[row] && typeof CELL_VALUE_CACHE[row][column] != "undefined") {
    return CELL_VALUE_CACHE[row][column];
  }

  if (!CELL_VALUE_CACHE[row]) {
    CELL_VALUE_CACHE[row] = [];
  }
  var range = sheet.getRange(row, column);
  return (CELL_VALUE_CACHE[row][column] = range && range.getValue() || emptyCellValue);
}

function getTeam() {
  var sheet = getStaticDataSheet();
  var endRow = sheet.createTextFinder("Team compositions").findNext().getRow() - 1;
  var team = {};
  for (var row = 4; row < endRow; ++row) {
    team[getStaticCellValue(sheet, row, 3).trim()] = {
      name: getStaticCellValue(sheet, row, 2).trim(),
      country: getStaticCellValue(sheet, row, 4).trim(),
      workWeek: getStaticCellValue(sheet, row, 5)
    };
  }
  addAvailabilityData(team)
  return team;
}

function getAllSprints() {
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
    value = getStaticCellValue(sheet, startRow, currentColumn).trim();
    if (value) {
      ++currentColumn;
      continue;
    }
    endColumn = currentColumn;
  }
  for (var column = startColumn; column < endColumn; ++column) {
    SPRINT_NAMES.push(getStaticCellValue(sheet, startRow, column).trim());
  }
  return SPRINT_NAMES;
}

function addAvailabilityData(team) {
  var sheet = getStaticDataSheet();
  var startRow = sheet.createTextFinder("PTO Data").findNext().getRow() + 2;
  var sprints = getAllSprints();
  var countryData = getAvailabilityPerCountry(team, sprints);

  var assigneeEmails = Object.keys(team);
  for (var member, factors, i = 0, l = assigneeEmails.length; i < l; ++i) {
    member = team[assigneeEmails[i]];
    member.availability = (member.workWeek / 40) * AVAILABILITY_BOOSTS.work_week;

    member.sprintsAvailability = {};
    for (var daysLeft, sprint, holidays, j = 0, l2 = sprints.length; j < l2; ++j) {
      daysLeft = SPRINT_LENGTH * member.availability;
      sprint = sprints[j];
      holidays = countryData[member.country] && countryData[member.country][sprint];
      if (holidays) {
        daysLeft -= holidays * AVAILABILITY_BOOSTS.holiday;
      }
      daysLeft -= (getStaticCellValue(sheet, startRow + i, j + 2) || 0) * AVAILABILITY_BOOSTS.pto;
      member.sprintsAvailability[sprint] = Math.max(daysLeft, 0) / SPRINT_LENGTH;
    }
  }

  return team;
}

function getAvailabilityPerCountry(team, sprints) {
  var sheet = getStaticDataSheet();
  
  var assigneeEmails = Object.keys(team);
  var countries = [];
  for (var member, i = 0, l = assigneeEmails.length; i < l; ++i) {
    country = team[assigneeEmails[i]].country;
    if (countries.indexOf(country) == -1) {
      countries.push(country);
    }
  }

  var startRow = sheet.createTextFinder("PTO Data").findNext().getRow() + Object.keys(team).length + 3;
  var endRow = startRow + countries.length;
  var countryData = {};
  for (var countryName, j = startRow; j < endRow; ++j) {
    countryName = getStaticCellValue(sheet, j, 1).trim();
    if (countries.indexOf(countryName) == -1) {
      continue;
    }
    countryData[countryName] = {};
    for (var k = 0, l2 = sprints.length; k < l2; ++k) {
      countryData[countryName][sprints[k]] = (getStaticCellValue(sheet, j, k + 2) || 0);
    }
  }
  return countryData;
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

function parseTeamsLine(line, sprintIndex, allTeamNames) {
  var teams = { active: [], commitments: [], pointsCommitted: {} };
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
      teams.pointsCommitted[teamName] = addSprintPointsCommitted(teamName, sprintIndex);
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
  var dataRow = sheet.createTextFinder("Team compositions").findNext().getRow() + 3;

  var sprints = getAllSprints();
  var sprintData = {};
  var allTeamNames = [];
  for (var line, i = 0, l = sprints.length; i < l; ++i) {
    line = getStaticCellValue(sheet, dataRow, i + 2).trim();
    if (!line) {
      continue;
    }
    sprintData[sprints[i]] = parseTeamsLine(line, i, allTeamNames);
    // addSprintPointsCommitted(sprintData[sprints[i]], i);
  }

  return {
    sprints: sprintData,
    names: allTeamNames.sort().sort(function(a, b) {
      return (b.match(/,/g) || []).length - (a.match(/,/g) || []).length;
    })
  };
}

function getForecastTeams() {
  var sheet = getStaticDataSheet();
  var dataRow = sheet.createTextFinder("Forecast teams").findNext().getRow();

  var teams = parseTeamsLine(getStaticCellValue(sheet, dataRow, 2).trim(), 0, []);
  return teams;
}

function getTeamCommitments(teams, assigneeNames) {
  for (var team, present, i = 0, l = teams.length; i < l; ++i) {
    team = teams[i];
    if (Object.keys(team).length != assigneeNames.length) {
      continue;
    }

    present = [];
    for (var j = 0, l2 = assigneeNames.length; j < l2; ++j) {
      if (!team[assigneeNames[j]]) {
        break;
      }
      present.push(assigneeNames[j]);
    }
    if (present.length == assigneeNames.length) {
      return team;
    }
  }

  // Return something ridiculous so that the table displays weird results.
  var fakeCommitments = {};
  for (var k = 0, l3 = assigneeNames.length; k < l3; ++k) {
    fakeCommitments[assigneeNames[k]] = 100000;
  }
  return fakeCommitments;
}

function addSprintPointsCommitted(teamName, sprintIndex) {
  var sheet = getStaticDataSheet();
  var startRange = sheet.createTextFinder("Sprint points committed").findNext();
  var range = sheet.createTextFinder(teamName).startFrom(startRange).findNext();
  if (!range) {
    return 0;
  }

  var teamRow = range.getRow();
  return getStaticCellValue(sheet, teamRow, sprintIndex + 2, 0);
}
