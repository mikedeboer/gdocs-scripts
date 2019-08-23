var ALPHA = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
var SPRINT_COLUMN_MAP = {};
var ASSIGNEE_ROW_MAP = {};

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

  var points = getPoints(root, Object.keys(team));
  var sheet = generateSheet(points, team);
  drawCharts(sheet, points, team);

  addTeamsSection(sheet, points, team);
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

function getPoints(root, assignees) {
  // The report lists the assignees inside the table with the `@domain.tld` email suffix,
  // So let's make a map to account for that efficiently.
  var assigneeMap = {};
  for (var assignee, i = 0, l = assignees.length; i < l; ++i) {
    assignee = assignees[i];
    assigneeMap[assignee.substr(0, assignee.indexOf("@"))] = assignee;
  }
  // Format of the data to return: { sprint: { assignee: points } }
  var pointsMap = {};
  var sprints = getSprints(root);
  for (var j = 0, l2 = sprints.length; j < l2; ++j) {
    pointsMap[sprints[j]] = {};
    for (var k = 0, l3 = assignees.length; k < l3; ++k) {
      pointsMap[sprints[j]][assignees[k]] = {};
    }
  }

  ["---", 1, 2, 3, 5, 8].forEach(function(points) {
    var container = findElements(root, {
      attrName: "id",
      attrValue: "tabular_report_container_" + points
    })[0];
    var rows = findElements(container, { tagName: "tr" });
    var currentAssignee, sprintIndex, currentSprint;
    // Cells: | assignee | --- | 65.4 - Dec 3-9 |  68.2 - Apr 1 - 14 | 68.3 - Apr 15 - 28 | Total |
    for (var row, cells, i = 0, l = rows.length; i < l; ++i) {
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
          currentAssignee = assigneeMap[cell];
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
        pointsMap[currentSprint][currentAssignee][points + ""] =
          (points == "---") ? cell : cell * points;
        ++sprintIndex;
      }
    }
  });

  return pointsMap;
}

function generateSheet(points, team) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.insertSheet("Velocity - generated at " + (new Date()).toDateString());
  var sprintNames = Object.keys(points);
  var sprintCount = sprintNames.length;
  var visitedRanges = {};
  for (var sprintName, assigneeEmails, assigneeCount, sprintColumn, i = 0; i < sprintCount; ++i) {
    sprintName = sprintNames[i];
    assigneeEmails = Object.keys(points[sprintName]);
    assigneeCount = assigneeEmails.length;
    sprintColumn = i + 2;

    if (!visitedRanges["1," + sprintColumn]) {
      visitedRanges["1," + sprintColumn] = 1;
      sheet.getRange(1, sprintColumn, 1, 1).setValue(sprintName);
      SPRINT_COLUMN_MAP[sprintName] = sprintColumn;
      var A1RangeSprintTotal = ALPHA[i + 1] + "2:" + ALPHA[i + 1] + (assigneeCount + 1);
      sheet.getRange(assigneeEmails.length + 3, sprintColumn, 1, 1).setValue("=SUM(" + A1RangeSprintTotal + ")");
    }

    for (var assignee, assigneeRow, pointTotal, note, pointTypes, j = 0; j < assigneeCount; ++j) {
      assignee = assigneeEmails[j];
      assigneeRow = j + 2;

      if (!visitedRanges[assigneeRow + ",1"]) {
        visitedRanges[assigneeRow + ",1"] = 1;
        sheet.getRange(assigneeRow, 1, 1, 1).setValue(team[assignee].name);
        ASSIGNEE_ROW_MAP[team[assignee].name] = assigneeRow;
        var A1RangeAssigneeTotal = "B" + assigneeRow + ":" + ALPHA[sprintCount] + assigneeRow;
        sheet.getRange(assigneeRow, sprintCount + 3, 1, 1).setValue("=SUM(" + A1RangeAssigneeTotal + ")");
      }

      pointTotal = 0;
      pointTypes = Object.keys(points[sprintName][assignee]);
      for (var pointValue, k = 0, l3 = pointTypes.length; k < l3; ++k) {
        pointValue = points[sprintName][assignee][pointTypes[k]];
        if (!pointValue) {
          continue;
        }
        pointTotal += pointValue;
      }
      sheet.getRange(assigneeRow, sprintColumn, 1, 1).setValue(pointTotal);
    }
  }
  return sheet;
}

function drawCharts(sheet, points, team) {
  var sprints = Object.keys(points);
  var assigneeEmails = Object.keys(team);
  // First, chart out the developments per sprint.
  var chart = sheet.newChart()
    .asLineChart()
    .addRange(sheet.getRange("B1:" + ALPHA[sprints.length] + "1"))
    .addRange(sheet.getRange("B" + (assigneeEmails.length + 3) + ":" + ALPHA[sprints.length] + (assigneeEmails.length + 3)))
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_ROWS)
    .setTransposeRowsAndColumns(true)
    .setXAxisTitle("Sprints")
    .setYAxisTitle("Points")
    .setPosition(assigneeEmails.length + 5, assigneeEmails.length + 5, -1000, 0)
    .build();
  sheet.insertChart(chart);

  // Next, chart out the developments per assignee.
  chart = sheet.newChart()
    .asPieChart()
    .addRange(sheet.getRange("A2:A" + (assigneeEmails.length + 1)))
    .addRange(sheet.getRange(ALPHA[sprints.length + 2] + "2:" + ALPHA[sprints.length + 2] + (assigneeEmails.length + 1)))
    .setPosition(assigneeEmails.length + 5, assigneeEmails.length + 5, -350, 0)
    .build();
  sheet.insertChart(chart);
}

function addTeamsSection(sheet, points, team) {
  var teamData = getTeamData();
  //Logger.log("TEAMS:: " + JSON.stringify(teamData, null, 2));

  var mainTableSpan = Object.keys(team).length + 4;
  var chartsRowSpan = 19;
  var tableRowOffset = mainTableSpan + chartsRowSpan;

  var teamCount = teamData.names.length;
  var sprintNames = Object.keys(points);
  var sprintCount = sprintNames.length;
  var visitedRanges = {};
  for (var range, sprintName, sprintColumn, i = 0; i < sprintCount; ++i) {
    sprintName = sprintNames[i];
    sprintColumn = i + 2;

    range = tableRowOffset + 1 + "," + sprintColumn;
    if (!visitedRanges[range]) {
      sheet.getRange(tableRowOffset + 1, sprintColumn, 1, 1).setValue(sprintName);
      var A1RangeSprintTotal = ALPHA[i + 1] + (tableRowOffset + 2) + ":" + ALPHA[i + 1] + (tableRowOffset + teamCount + 1);
      sheet.getRange(tableRowOffset + teamCount + 3, sprintColumn, 1, 1).setValue("=SUM(" + A1RangeSprintTotal + ")");
      visitedRanges[range] = 1;
    }

    for (var teamName, teamPoints, teamRow, assignees, teamCommitments, j = 0; j < teamCount; ++j) {
      teamName = teamData.names[j];
      teamPoints = 0;
      teamRow = tableRowOffset + j + 2;
      // If this team was active during this sprint, get its relative velocity.
      if (teamData.sprints[sprintName] && teamData.sprints[sprintName].active.indexOf(teamName) != -1) {
        if (!visitedRanges[teamRow + ",1"]) {
          sheet.getRange(teamRow, 1, 1, 1).setValue(teamName);
          var A1RangeTeamTotal = "B" + teamRow + ":" + ALPHA[sprintCount] + teamRow;
          sheet.getRange(teamRow, sprintCount + 3, 1, 1).setValue("=SUM(" + A1RangeTeamTotal + ")");
          visitedRanges[teamRow + ",1"] = 1;
        }

        assignees = teamName.split(/\s*,\s*/);
        teamCommitments = getTeamCommitments(teamData.sprints[sprintName].commitments, assignees);
        for (var assignee, commitment, points, k = 0, l = assignees.length; k < l; ++k) {
          assignee = assignees[k];
          commitment = teamCommitments[assignee] / 100;
          points = parseInt(sheet.getRange(ASSIGNEE_ROW_MAP[assignee], SPRINT_COLUMN_MAP[sprintName], 1, 1).getValue(), 10);
          // Logger.log("Adding points to team " + teamName + " for sprint " +
          //   sprintName + " obo " + assignee + ": " + points + "x" + commitment);
          teamPoints += (points * commitment);
        }
      }

      sheet.getRange(teamRow, sprintColumn, 1, 1).setValue(teamPoints);
    }
  }
}
