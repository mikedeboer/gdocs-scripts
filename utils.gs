function findElements(element, query) {
  var isRegexValue = !!query.attrValue && (query.attrValue instanceof RegExp);
  var descendants = element.getDescendants();  
  var found = [];
  for (var attr, val, i = 0, l = descendants.length; i < l; ++i) {
    var elt = descendants[i].asElement();
    if (!elt) {
      continue;
    }
    
    if (query.tagName && elt.getName() != query.tagName) {
      continue;
    }
    
    if (!query.attrName) {
      if (query.tagName) {
        found.push(elt);
      }
      continue;
    }

    attr = elt.getAttribute(query.attrName);
    if (!attr) {
      continue;
    }

    if (!query.attrValue) {
      found.push(elt);
    }

    val = attr.getValue();
    if (isRegexValue ? query.attrValue.test(val) : val == query.attrValue) {
      found.push(elt);
    }
  }
  return found;
}

function sanitizeBugzillaHTML(html) {
  return html
    .match(/<main[^>]*>([\s\S]*)<\/main>/gim)[0] // Only keep the main body.
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gim, "") // Get rid of script tags.
    .replace(/<br[^\/>]*>/gim, "<br/>");
}

var ALPHA = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
var ALPHA_LEN = ALPHA.length;

function getR1C1(c1, r1, c2, r2) {
  --c1;
  var columnPrefix = ALPHA[Math.floor(c1 / ALPHA_LEN) - 1] || "";
  var r1c1 = columnPrefix + ALPHA[(c1 % ALPHA_LEN)] + r1;
  if (c2 && r2) {
    --c2;
    columnPrefix = ALPHA[Math.floor(c2 / ALPHA_LEN) - 1] || "";
    r1c1 += ":" + columnPrefix + ALPHA[(c2 % ALPHA_LEN)] + r2;
  }
  return r1c1;
}

function getAssigneeEmail(assigneeName, team) {
  var memberEmails = Object.keys(team);
  for (var i = 0, l = memberEmails.length; i < l; ++i) {
    if (team[memberEmails[i]].name == assigneeName) {
      return memberEmails[i];
    }
  }
  return null;
}

function getAvailability(name, sprint, team) {
  var memberEmails = Object.keys(team);
  for (var memberEmail, i = 0, l = memberEmails.length; i < l; ++i) {
    memberEmail = memberEmails[i];
    if (team[memberEmail].name != name) {
      continue;
    }
    return team[memberEmail].sprintsAvailability[sprint] || 1;
  }
  return 1;
}
