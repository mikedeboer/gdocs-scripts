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
