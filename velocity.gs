var ALPHA = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
var COLUMN_LABELS = {
  absolute: "Absolute",
  average: "Average",
  averageMinus: "Average minus STDEV",
  averagePlus: "Average plus STDEV",
  points: "Points",
  relative: "Relative",
  sprints: "Sprints",
  teams: "Teams",
  total: "Total",
  totalMinus: "Total minus STDEV",
  totalPlus: "Total plus STDEV"
};
var CHART_ROWSPAN = 19;

function generateVelocity() {
  var team = getTeam();
  var url = "https://bugzilla.mozilla.org/report.cgi?"
    + ["x_axis_field=cf_fx_iteration", "y_axis_field=assigned_to", "z_axis_field=cf_fx_points",
       "query_format=report-table", "short_desc_type=notregexp", "short_desc=%5E%5C%5Bmeta",
       "product=Core", "product=Firefox", "product=Toolkit",
       "bug_status=RESOLVED", "bug_status=VERIFIED", "bug_status=CLOSED", "resolution=FIXED",
       "longdesc_type=allwordssubstr", "longdesc=",
       "bug_file_loc_type=allwordssubstr", "bug_file_loc=",
       "status_whiteboard_type=allwordssubstr", "status_whiteboard=",
       "keywords_type=allwords", "keywords=",
       "bug_id=", "bug_id_type=anyexact",
       "votes=", "votes_type=greaterthaneq",
       "priority=P1", "priority=P2", "priority=P3", "priority=P4", "priority=P5", "priority=--",
       "emailtype1=notequals", "email1=intermittent-bug-filer%40mozilla.bugs",
       "emailtype2=substring", "email2=&emailtype3=substring",
       "email3=",
       "chfield=%5BBug+creation%5D", "chfieldvalue=", "chfieldfrom=", "chfieldto=",
       "j_top=AND", "f1=component", "o1=nowordssubstr", "v1=Graphics%3A+WebRender%2C+CA+Certifica%2C+Build+Config",
       "f2=keywords", "o2=notsubstring", "v2=meta",
       "f3=assigned_to", "o3=anyexact", "v3=" + encodeURIComponent(Object.keys(team).join(",")),
       "f4=cf_last_resolved", "o4=greaterthan", "v4=2019-04-14",
       "format=table", "action=wrap"
      ].join("&");
  var html = sanitizeBugzillaHTML(UrlFetchApp.fetch(url).getContentText());
  var root = XmlService.parse(html).getRootElement();

  var points = getPoints(root, team);
  var teamData = getTeamData();
  generateSheet(points, team, teamData);
}

function getSprints(root) {
  var container = findElements(root, {
    attrName: "id",
    attrValue: "tabular_report_container_---"
  })[0];
  var headers = findElements(findElements(container, { tagName: "thead" })[0], {
    tagName: "th",
    attrName: "class",
    attrValue: /^t\d$/
  });
  var sprints = [];
  for (var header, i = 0, l = headers.length; i < l; ++i) {
    header = (headers[i].getValue() || "").trim();
    if (!header || header.indexOf("--") > -1) {
      continue;
    }
    sprints.push(header);
  }
  return sprints;
}

function getPoints(root, team) {
  // The report lists the assignees inside the table with the `@domain.tld` email suffix,
  // So let's make a map to account for that efficiently.
  var assigneeMap = {};
  var assigneeEmails = Object.keys(team);
  for (var assigneeEmail, i = 0, l = assigneeEmails.length; i < l; ++i) {
    assigneeEmail = assigneeEmails[i];
    assigneeMap[assigneeEmail.substr(0, assigneeEmail.indexOf("@"))] = assigneeEmail;
  }
  // Format of the data to return: { sprint: { assignee: points } }
  var pointsMap = {};
  var sprints = getSprints(root);
  for (var j = 0, l2 = sprints.length; j < l2; ++j) {
    pointsMap[sprints[j]] = {};
    for (var k = 0, l3 = assigneeEmails.length; k < l3; ++k) {
      pointsMap[sprints[j]][assigneeEmails[k]] = {};
    }
  }

  ["---", 1, 2, 3, 5, 8].forEach(function(pointsWeight) {
    var container = findElements(root, {
      attrName: "id",
      attrValue: "tabular_report_container_" + pointsWeight
    })[0];
    var rows = findElements(container, { tagName: "tr" });
    var assigneeEmail, assigneeName, sprintIndex, currentSprint;
    // Cells: | assignee | --- | 65.4 - Dec 3-9 |  68.2 - Apr 1 - 14 | 68.3 - Apr 15 - 28 | Total |
    for (var row, cells, points, i = 0, l = rows.length; i < l; ++i) {
      row = rows[i];
      //Logger.log("ROW CONTENTS:: " + XmlService.getPrettyFormat().format(row));
      cells = findElements(row, { tagName: "td" });
      sprintIndex = 0;
      for (var cell, j = 0, l2 = cells.length; j < l2; ++j) {
        cell = (cells[j].getValue() || "").trim();
        currentSprint = sprints[sprintIndex];

        if (cell == "Total") {
          // Don't process 'Total' rows.
          break;
        } else if (assigneeMap[cell]) {
          // Found assignee, activate it.
          assigneeEmail = assigneeMap[cell];
          assigneeName = team[assigneeEmail].name;
          // Skip the next cell.
          ++j;
          continue;
        } else if (cell == ".") {
          ++sprintIndex;
          continue;
        } else if (!currentSprint) {
          continue;
        }
        cell = parseInt(cell, 10);
        if (pointsWeight == "---") {
          points = cell;
        } else {
          points = cell * pointsWeight;
          if (!pointsMap[currentSprint][assigneeEmail].total) {
            pointsMap[currentSprint][assigneeEmail].total = {
              absolute: 0,
              availabilityQuotient: (1 + (1 - getAvailability(assigneeName, currentSprint, team)))
            };
          }
          pointsMap[currentSprint][assigneeEmail].total.absolute += points;
        }
        pointsMap[currentSprint][assigneeEmail][pointsWeight + ""] = points;
        ++sprintIndex;
      }
    }
  });

  return pointsMap;
}

function generateSheet(points, team, teamData) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = "Velocity - generated at " + (new Date()).toDateString();
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (sheet) {
    spreadsheet.deleteSheet(sheet);
  }
  sheet = spreadsheet.insertSheet(sheetName);

  var assigneeEmailsCount = Object.keys(team).length;
  var teamCount = teamData.names.length;
  var startRow = 2;
  
  var sprintNames = Object.keys(points);
  var sprintCount = sprintNames.length;
  var totalColumn = sprintCount + 4;
  // Propagate the first row with the column-labels used throughout the sheet.
  for (var sprintName, sprintColumn, i = 0; i < sprintCount; ++i) {
    sprintColumn = i + 2;
    sprintName = sprintNames[i];
    sheet.getRange(1, sprintColumn, 1, 1).setValue(sprintName);
    sheet.getRange(1, totalColumn - 1, 1, 1).setValue(COLUMN_LABELS.totalMinus);
    sheet.getRange(1, totalColumn, 1, 1).setValue(COLUMN_LABELS.total);
    sheet.getRange(1, totalColumn + 1, 1, 1).setValue(COLUMN_LABELS.totalPlus);
    sheet.getRange(1, totalColumn + 2, 1, 1).setValue(COLUMN_LABELS.averageMinus);
    sheet.getRange(1, totalColumn + 3, 1, 1).setValue(COLUMN_LABELS.average);
    sheet.getRange(1, totalColumn + 4, 1, 1).setValue(COLUMN_LABELS.averagePlus);
  }
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
  sheet.hideColumns(totalColumn - 1);
  sheet.hideColumns(totalColumn + 1);
  sheet.hideColumns(totalColumn + 2);
  sheet.hideColumns(totalColumn + 4);

  var rowSpan = addForecastsSection(sheet, points, team, teamData, startRow);
  rowSpan += addTeamsSection(sheet, points, team, teamData, startRow + rowSpan);
  rowSpan += addIndividualsSection(sheet, points, team, teamData, startRow + rowSpan);
}

function addIndividualsSection(sheet, points, team, teamData, startRow) {
  var sprintNames = Object.keys(points);
  var sprintCount = sprintNames.length;
  var visitedRanges = {};

  var tableStartRow = startRow + CHART_ROWSPAN;
  var sprintName, assigneeEmails, assigneeCount, relativeVelocityTableStartRow,
    sprintColumn, totalRow;
  for (var i = 0; i < sprintCount; ++i) {
    sprintName = sprintNames[i];
    assigneeEmails = Object.keys(points[sprintName]);
    assigneeCount = assigneeEmails.length;
    relativeVelocityTableStartRow = tableStartRow + assigneeCount + 5;
    sprintColumn = i + 2;

    if (!visitedRanges[tableStartRow + "," + sprintColumn]) {
      visitedRanges[tableStartRow + "," + sprintColumn] = 1;
      totalRow = tableStartRow + assigneeCount + 2;
      sheet.getRange(totalRow - 1, 1, 1, 1).setValue(COLUMN_LABELS.totalMinus);
      sheet.getRange(totalRow, 1, 1, 1).setValue(COLUMN_LABELS.total);
      sheet.getRange(totalRow + 1, 1, 1, 1).setValue(COLUMN_LABELS.totalPlus);
      var A1RangeSprintTotal = getR1C1(i + 2, tableStartRow, i + 2, tableStartRow + assigneeCount - 1);
      sheet.getRange(totalRow - 1, sprintColumn, 1, 1).setValue("=" +
        getR1C1(sprintColumn, totalRow) + "-STDEV(" + A1RangeSprintTotal + ")");
      sheet.getRange(totalRow, sprintColumn, 1, 1).setValue("=SUM(" + A1RangeSprintTotal + ")");
      sheet.getRange(totalRow + 1, sprintColumn, 1, 1).setValue("=" +
        getR1C1(sprintColumn, totalRow) + "+STDEV(" + A1RangeSprintTotal + ")");
      sheet.hideRows(totalRow - 1);
      sheet.hideRows(totalRow + 1);

      // Also set the labels for the table with relative velocity (sprint totals).
      totalRow = relativeVelocityTableStartRow + assigneeCount + 2;
      sheet.getRange(totalRow - 1, 1, 1, 1).setValue(COLUMN_LABELS.totalMinus);
      sheet.getRange(totalRow, 1, 1, 1).setValue(COLUMN_LABELS.total);
      sheet.getRange(totalRow + 1, 1, 1, 1).setValue(COLUMN_LABELS.totalPlus);
      A1RangeSprintTotal = getR1C1(i + 2, relativeVelocityTableStartRow, i + 2,
        relativeVelocityTableStartRow + assigneeCount - 1);
      sheet.getRange(totalRow - 1, sprintColumn, 1, 1).setValue("=" +
        getR1C1(sprintColumn, totalRow) + "-STDEV(" + A1RangeSprintTotal + ")");
      sheet.getRange(totalRow, sprintColumn, 1, 1).setValue("=SUM(" + A1RangeSprintTotal + ")");
      sheet.getRange(totalRow + 1, sprintColumn, 1, 1).setValue("=" +
        getR1C1(sprintColumn, totalRow) + "+STDEV(" + A1RangeSprintTotal + ")");
      sheet.hideRows(totalRow - 1);
      sheet.hideRows(totalRow + 1);
    }

    var assigneeEmail, assigneeName, pointTotal, note, pointTypes;
    for (var j = 0; j < assigneeCount; ++j) {
      assigneeEmail = assigneeEmails[j];
      assigneeName = team[assigneeEmail].name;

      if (!visitedRanges[tableStartRow + j + ",1"]) {
        totalRow = tableStartRow + j;
        visitedRanges[totalRow + ",1"] = 1;
        sheet.getRange(totalRow, 1, 1, 1).setValue(assigneeName);
        var totalColumn = sprintCount + 4;
        var A1RangeAssigneeTotal = getR1C1(2, totalRow, sprintCount + 1, totalRow);
        sheet.getRange(totalRow, totalColumn - 1, 1, 1).setValue("=" +
          getR1C1(totalColumn, totalRow) + "-STDEV(" + A1RangeAssigneeTotal + ")");
        sheet.getRange(totalRow, totalColumn, 1, 1).setValue("=SUM(" + A1RangeAssigneeTotal + ")");
        sheet.getRange(totalRow, totalColumn + 1, 1, 1).setValue("=" +
          getR1C1(totalColumn, totalRow) + "+STDEV(" + A1RangeAssigneeTotal + ")");

        // Also set the labels for the table with relative velocity (assignee totals).
        totalRow = relativeVelocityTableStartRow + j;
        sheet.getRange(totalRow, 1, 1, 1).setValue(assigneeName);
        A1RangeAssigneeTotal = getR1C1(2, totalRow, sprintCount + 1, totalRow);
        sheet.getRange(totalRow, totalColumn - 1, 1, 1).setValue("=" +
          getR1C1(totalColumn, totalRow) + "-STDEV(" + A1RangeAssigneeTotal + ")");
        sheet.getRange(totalRow, totalColumn, 1, 1).setValue("=SUM(" + A1RangeAssigneeTotal + ")");
        sheet.getRange(totalRow, totalColumn + 1, 1, 1).setValue("=" +
          getR1C1(totalColumn, totalRow) + "+STDEV(" + A1RangeAssigneeTotal + ")");
      }

      var pointsTotal = points[sprintName][assigneeEmail].total || {
        absolute: 0,
        availabilityQuotient: 1
      };
      sheet.getRange(tableStartRow + j, sprintColumn, 1, 1).setValue(pointsTotal.absolute);
      // Also set the values for the table with relative velocity.
      sheet.getRange(relativeVelocityTableStartRow + j, sprintColumn, 1, 1)
        .setValue(pointsTotal.absolute * pointsTotal.availabilityQuotient);
    }
  }

  drawIndividualsCharts(sheet, points, team, teamData, tableStartRow);
  return CHART_ROWSPAN + (assigneeCount * 2) + 12;
}

function drawIndividualsCharts(sheet, points, team, teamData, tableStartRow) {
  var sprintCount = Object.keys(points).length;
  var assigneeEmailsCount = Object.keys(team).length;
  var chartsOffset = 5;
  // First, chart out the developments per sprint.
  var totalRow = tableStartRow + (assigneeEmailsCount * 2) + 7;
  var totalColumn = sprintCount + 4;

  var chart = sheet.newChart()
    .asAreaChart()
    .addRange(sheet.getRange(getR1C1(2, 1, sprintCount + 1, 1)))
    .addRange(sheet.getRange(getR1C1(1, totalRow, sprintCount + 1, totalRow)))
    .addRange(sheet.getRange(getR1C1(1, totalRow, sprintCount + 1, totalRow - 1)))
    .addRange(sheet.getRange(getR1C1(1, totalRow + 1, sprintCount + 1, totalRow + 1)))
    .setNumHeaders(1)
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_ROWS)
    .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.SHOW_BOTH)
    .setTransposeRowsAndColumns(true)
    .setXAxisTitle(COLUMN_LABELS.sprints)
    .setYAxisTitle(COLUMN_LABELS.points)
    .setPosition(tableStartRow - (assigneeEmailsCount * 2) - chartsOffset, 2, 0, 0);
  sheet.insertChart(chart.build());

  // Next, chart out the developments per assignee.
  chart = sheet.newChart()
    .asPieChart()
    .addRange(sheet.getRange(getR1C1(1, tableStartRow, 1, tableStartRow + assigneeEmailsCount - 1)))
    .addRange(sheet.getRange(getR1C1(totalColumn, tableStartRow, totalColumn,
      tableStartRow + assigneeEmailsCount - 1)))
    .setPosition(tableStartRow - (assigneeEmailsCount * 2) - chartsOffset, 9, 0, 0);
  sheet.insertChart(chart.build());

  return sheet;
}

function addTeamsSection(sheet, points, team, teamData, startRow) {
  var teamCount = teamData.names.length;
  var tableRowOffset = startRow + CHART_ROWSPAN + teamCount;

  var sprintNames = Object.keys(points);
  var sprintCount = sprintNames.length;
  var visitedRanges = {};
  var teamSprintData = {};
  var teamsCommittedCount = 0;
  var range, sprintName, sprintColumn;
  for (var i = 0; i < sprintCount; ++i) {
    sprintName = sprintNames[i];
    sprintColumn = i + 2;

    range = tableRowOffset + "," + sprintColumn;
    if (!visitedRanges[range]) {
      var totalRow = tableRowOffset + teamCount + 3;
      sheet.getRange(totalRow - 1, 1, 1, 1).setValue(COLUMN_LABELS.totalMinus);
      sheet.getRange(totalRow, 1, 1, 1).setValue(COLUMN_LABELS.total);
      sheet.getRange(totalRow + 1, 1, 1, 1).setValue(COLUMN_LABELS.totalPlus);
      var A1RangeSprintTotal = getR1C1(i + 2, tableRowOffset + 2, i + 2,
        tableRowOffset + teamCount + 1);
      sheet.getRange(totalRow - 1, sprintColumn, 1, 1).setValue("=" +
        getR1C1(sprintColumn, totalRow) + "-STDEV(" + A1RangeSprintTotal + ")");
      sheet.getRange(totalRow, sprintColumn, 1, 1).setValue("=SUM(" + A1RangeSprintTotal + ")");
      sheet.getRange(totalRow + 1, sprintColumn, 1, 1).setValue("=" +
        getR1C1(sprintColumn, totalRow) + "+STDEV(" + A1RangeSprintTotal + ")");
      sheet.hideRows(totalRow - 1);
      sheet.hideRows(totalRow + 1);
      visitedRanges[range] = 1;
    }

    var teamName, teamPoints, teamRow, assigneeNames, teamCommitments;
    for (var j = 0; j < teamCount; ++j) {
      teamName = teamData.names[j];
      teamPoints = 0;
      teamRow = tableRowOffset + j + 1;
      // If this team was active during this sprint, get its relative velocity.
      if (teamData.sprints[sprintName] && teamData.sprints[sprintName].active.indexOf(teamName) != -1) {
        if (!teamSprintData[teamName]) {
          teamSprintData[teamName] = [];
        }
        if (!visitedRanges[teamRow + ",1"]) {
          sheet.getRange(teamRow, 1, 1, 1).setValue(teamName);
          var totalColumn = sprintCount + 4;
          var A1RangeTeamTotal = getR1C1(2, teamRow, sprintCount + 1, teamRow);
          sheet.getRange(teamRow, totalColumn - 1, 1, 1).setValue("=" +
            getR1C1(totalColumn, teamRow) + "-STDEV(" + A1RangeTeamTotal + ")");
          sheet.getRange(teamRow, totalColumn, 1, 1).setValue("=SUM(" + A1RangeTeamTotal + ")");
          sheet.getRange(teamRow, totalColumn + 1, 1, 1).setValue("=" +
            getR1C1(totalColumn, teamRow) + "+STDEV(" + A1RangeTeamTotal + ")");
          visitedRanges[teamRow + ",1"] = 1;

          // Also add labels for the calculations table.
          sheet.getRange(teamRow - teamCount - 1, 1, 1, 1).setValue(teamName);
        }

        assigneeNames = teamName.split(/\s*,\s*/);
        teamCommitments = getTeamCommitments(teamData.sprints[sprintName].commitments, assigneeNames);
        var assigneeName, assigneeAvailability, commitment, assigneePoints;
        for (var k = 0, l = assigneeNames.length; k < l; ++k) {
          assigneeName = assigneeNames[k];
          assigneeAvailability = 1 + (1 - getAvailability(assigneeName, sprintName, team));
          commitment = teamCommitments[assigneeName] / 100;
          pointsTotal = (points[sprintName][getAssigneeEmail(assigneeName, team)] || {}).total;
          if (!pointsTotal) {
            Logger.log("No data found for " + sprintName + ":" + assigneeName);
            continue;
          }

          assigneePoints = pointsTotal.absolute;
          teamPoints += ((pointsTotal.absolute * assigneeAvailability) * commitment);
        }
      }

      sheet.getRange(teamRow, sprintColumn, 1, 1).setValue(teamPoints);
      if (teamPoints) {
        teamSprintData[teamName].push({
          points: teamPoints,
          sprintName: sprintName,
          pointsCommitted: teamData.sprints[sprintName].pointsCommitted[teamName] || 0
        });
      }
    }
  }

  // Add totals to the calculations table.
  var pointsCount;
  for (p = 0; p < teamCount; ++p) {
    teamRow = (tableRowOffset - teamCount) + p;
    teamName = teamData.names[p];
    teamPoints = pruneOutliers(teamSprintData[teamName].map(function(o) {
      return o.points;
    }));
    pointsCount = teamPoints.length;
    for (var q = 0; q < pointsCount; ++q) {
      sheet.getRange(teamRow, q + 2, 1, 1).setValue(teamPoints[q]);
    }
    A1RangeTeamTotal = getR1C1(2, teamRow, pointsCount, teamRow);
    sheet.getRange(teamRow, totalColumn).setValue("=SUM(" + A1RangeTeamTotal + ")");
    sheet.getRange(teamRow, totalColumn + 3).setValue("=AVERAGE(" + A1RangeTeamTotal + ")");
    if (pointsCount > 2) {
      // Totals;
      sheet.getRange(teamRow, totalColumn - 1).setValue("=" +
        getR1C1(totalColumn, teamRow) + "-STDEV(" + A1RangeTeamTotal + ")");
      sheet.getRange(teamRow, totalColumn + 1).setValue("=" +
        getR1C1(totalColumn, teamRow) + "+STDEV(" + A1RangeTeamTotal + ")");
      // Averages;
      sheet.getRange(teamRow, totalColumn + 2).setValue("=STDEV(" + A1RangeTeamTotal + ")");
      sheet.getRange(teamRow, totalColumn + 4).setValue("=STDEV(" + A1RangeTeamTotal + ")");
    } else {
      // Totals;
      sheet.getRange(teamRow, totalColumn - 1).setValue("=SUM(" + A1RangeTeamTotal + ")");
      sheet.getRange(teamRow, totalColumn + 1).setValue("=SUM(" + A1RangeTeamTotal + ")");
      // Averages;
      sheet.getRange(teamRow, totalColumn + 2).setValue(0);
      sheet.getRange(teamRow, totalColumn + 4).setValue(0);
    }

    // Now check the points that we registered with the amount that the team
    // committed to, during the sprint planning (before the sprint started).
    teamPoints = teamSprintData[teamName].filter(function(o) {
      return !!o.pointsCommitted && !!o.points;
    });
    pointsCount = teamPoints.length;
    if (!pointsCount) {
      continue;
    }
    ++teamsCommittedCount
    var estimationRating;
    for (i = 0; i < pointsCount; ++i) {
      teamRow = tableRowOffset + teamCount + teamsCommittedCount + 5;
      sprintColumn = sprintNames.indexOf(teamPoints[i].sprintName) + 2;
      range = teamRow + ",1";
      if (!visitedRanges[range]) {
        sheet.getRange(teamRow, 1).setValue(teamName);
      }
      Logger.log("CALC:: " + teamPoints[i].points + ", " + teamPoints[i].pointsCommitted);
      estimationRating = ((teamPoints[i].points - teamPoints[i].pointsCommitted)
        / teamPoints[i].pointsCommitted) * 100;
      sheet.getRange(teamRow, sprintColumn).setValue(estimationRating.toFixed(2));
    }
  }

  addTeamCharts(sheet, points, team, teamData, startRow);
  return CHART_ROWSPAN + (teamCount * 2) + teamsCommittedCount + 7;
}

function addTeamCharts(sheet, points, team, teamData, startRow) {
  var sprintCount = Object.keys(points).length;
  var assigneeEmailsCount = Object.keys(team).length;
  var teamCount = teamData.names.length;

  var dataStartRow = startRow + CHART_ROWSPAN;
  var dataEndRow = dataStartRow + teamCount - 1;
  var averageColumn = sprintCount + 6;
  var chart = sheet.newChart()
    // .asAreaChart()
    .asColumnChart()
    .setStacked()
    .addRange(sheet.getRange(getR1C1(1, dataStartRow, 1, dataEndRow)))
    .addRange(sheet.getRange(getR1C1(averageColumn, dataStartRow, averageColumn, dataEndRow)))
    .addRange(sheet.getRange(getR1C1(averageColumn + 1, dataStartRow, averageColumn + 1, dataEndRow)))
    .addRange(sheet.getRange(getR1C1(averageColumn + 2, dataStartRow, averageColumn + 2, dataEndRow)))
    .setNumHeaders(1)
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.SHOW_BOTH)
    .setXAxisTitle(COLUMN_LABELS.teams)
    .setYAxisTitle(COLUMN_LABELS.points)
    .setPosition(startRow, 2, 0, 0);
  sheet.insertChart(chart.build());

  return sheet;
}

function addForecastsSection(sheet, points, team, teamData, startRow) {
  var forecastTeams = getForecastTeams();
  var forecastTeamCount = forecastTeams.active.length;
  var sprintNames = Object.keys(points);
  var allSprints = getAllSprints();

  var firstSprintIdx = allSprints.length - 1;
  for (; firstSprintIdx >= 0; --firstSprintIdx) {
    if (sprintNames.indexOf(allSprints[firstSprintIdx]) > -1) {
      break;
    }
  }
  var futureSprints = allSprints.slice(firstSprintIdx);
  var futureSprintCount = futureSprints.length;

  // First we need to get the average velocity numbers for independent assignees.
  var sprintCount = sprintNames.length;
  var assigneeTotals = {};
  var sprintName, assigneeEmails, assigneeCount;
  for (var i = 0; i < sprintCount; ++i) {
    sprintName = sprintNames[i];
    assigneeEmails = Object.keys(points[sprintName]);
    assigneeCount = assigneeEmails.length;

    var assigneeName, assigneeEmail, assigneeTotal;
    for (var j = 0; j < assigneeCount; ++j) {
      assigneeEmail = assigneeEmails[j];
      assigneeName = team[assigneeEmail].name;
      if (!assigneeTotals[assigneeName]) {
        assigneeTotals[assigneeName] = {
          absolute: [],
          relative: []
        };
      }
      assigneeTotal = (points[sprintName][assigneeEmail] || {}).total || {
        absolute: 0,
        availabilityQuotient: 1,
      };
      assigneeTotals[assigneeName].absolute.push(assigneeTotal.absolute);
      assigneeTotals[assigneeName].relative.push(assigneeTotal.absolute / assigneeTotal.availabilityQuotient);
    }
  }

  var assigneeAverages = {};
  var totalsCount, prunedAbsolute, prunedRelative;
  var assigneeNames = Object.keys(assigneeTotals);
  assigneeCount = assigneeNames.length;
  for (i = 0; i < assigneeCount; ++i) {
    assigneeName = assigneeNames[i];
    totalsCount = assigneeTotals[assigneeName].absolute.length;
    prunedAbsolute = pruneOutliers(assigneeTotals[assigneeName].absolute);
    prunedRelative = pruneOutliers(assigneeTotals[assigneeName].relative);
    assigneeAverages[assigneeName] = {
      absolute: {
        mean: mean(prunedAbsolute),
        stdev: standardDeviation(prunedAbsolute)
      },
      relative: {
        mean: mean(prunedRelative),
        stdev: standardDeviation(prunedRelative)
      }
    };
  }

  var futurePoints = {};
  var tableRowOffset = startRow + CHART_ROWSPAN;
  var sprintColumn;
  for (i = 0; i < futureSprintCount; ++i) {
    sprintName = futureSprints[i];
    sprintColumn = i + 2;
    sheet.getRange(tableRowOffset, sprintColumn).setValue(sprintName);

    futurePoints[sprintName] = {};
    var teamName, teamPoints, teamCommitments, teamRow;
    for (j = 0; j < forecastTeamCount; ++j) {
      teamName = forecastTeams.active[j];
      teamRow = tableRowOffset + (j * 2);
      sheet.getRange(teamRow + 1, 1).setValue(teamName + " - " + COLUMN_LABELS.absolute);
      sheet.getRange(teamRow + 2, 1).setValue(teamName + " - " + COLUMN_LABELS.relative);
      sheet.hideRows(teamRow + 2);

      teamPoints = futurePoints[sprintName][teamName] = {
        absolute: 0,
        relative: 0
      };
      assigneeNames = teamName.split(/\s*,\s*/);
      teamCommitments = getTeamCommitments(forecastTeams.commitments, assigneeNames);

      var assigneeEmail, assigneeAvailability, commitment, assigneePoints, pointsTillNow;
      for (var k = 0, l2 = assigneeNames.length; k < l2; ++k) {
        assigneeName = assigneeNames[k];
        assigneeEmail = getAssigneeEmail(assigneeName, team);
        assigneeAvailability = 1 + (1 - getAvailability(assigneeName, sprintName, team));
        commitment = teamCommitments[assigneeName] / 100;
        // First capture the points that are already handled this sprint.
        pointsTillNow = (points[sprintName] && (points[sprintName][assigneeEmail] || {}).total) || {
          absolute: 0,
          availabilityQuotient: 1
        };

        assigneePoints = assigneeAverages[assigneeName];
        teamPoints.absolute += (((assigneePoints.absolute.mean / assigneeAvailability) / commitment)
          - pointsTillNow.absolute);
        teamPoints.relative += (((assigneePoints.relative.mean / assigneeAvailability) / commitment)
          - pointsTillNow.absolute * pointsTillNow.availabilityQuotient);

        sheet.getRange(teamRow + 1, sprintColumn).setValue(teamPoints.absolute);
        sheet.getRange(teamRow + 2, sprintColumn).setValue(teamPoints.relative);
      }
    }
  }

  addForecastCharts(sheet, futurePoints, forecastTeams, startRow);

  return CHART_ROWSPAN + (forecastTeamCount * 2) + 2;
}

function addForecastCharts(sheet, futurePoints, forecastTeams, startRow) {
  var sprintCount = Object.keys(futurePoints).length;
  var teamCount = forecastTeams.active.length;
  var tableRowOffset = startRow + CHART_ROWSPAN;

  for (var chart, teamRow, i = 0; i < teamCount; ++i) {
    teamRow = tableRowOffset + (i * 2);
    chart = sheet.newChart()
      .asAreaChart()
      .addRange(sheet.getRange(getR1C1(2, tableRowOffset, sprintCount + 1, tableRowOffset)))
      .addRange(sheet.getRange(getR1C1(1, teamRow + 1, sprintCount + 1, teamRow + 1)))
      .addRange(sheet.getRange(getR1C1(1, teamRow + 2, sprintCount + 1, teamRow + 2)))
      .setNumHeaders(1)
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_ROWS)
      .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.SHOW_BOTH)
      .setTransposeRowsAndColumns(true)
      .setXAxisTitle(COLUMN_LABELS.sprints)
      .setYAxisTitle(COLUMN_LABELS.points)
      .setPosition(startRow, (i * 7) + 2, 0, 0);
    sheet.insertChart(chart.build());
  }
}
