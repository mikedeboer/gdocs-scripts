#!/usr/bin/env node
"use strict";

const css = require("css");
const fs = require("fs").promises;
const https = require("https");
const path = require("path");

const program = require("commander");
program.version("0.0.1");

const rootPath = path.normalize(path.join(__dirname, ".."));
const buildPath = path.join(rootPath, "tools", "build");
const hljsPath = path.join(rootPath, "tools", "node_modules", "highlight.js");
const hljsCDN = "https://cdnjs.cloudflare.com/ajax/libs/highlight.js/9.18.1/highlight.min.js";
const hljsFilename = "highlight.min.js";
const stylesMapFilename = "hljsStyles.js";
const hljsBundleFilename = "highlight.bundle.min.js";
const skipStyles = new Set([
  "brown-paper.css", "pojoaque.css", "school-book.css"
]);

function camelize(s) {
  s = s.split("-").map(p => p.charAt(0).toUpperCase() + p.substr(1)).join("");
  return s.charAt(0).toLowerCase() + s.substr(1);
}

async function clean() {
  try {
    await fs.unlink(path.join(buildPath, hljsFilename));
    await fs.unlink(path.join(buildPath, stylesMapFilename));
    await fs.unlink(path.join(buildPath, hljsBundleFilename));
  } catch (ex) {
    console.error("Error whilst cleaning up: ", ex);
  }
}

async function downloadScript() {
  let script;
  try {
    script = await fs.readFile(path.join(buildPath, hljsFilename), "utf8");
    if (script) {
      return script;
    }
  } catch (ex) {}

  script = await new Promise((resolve, reject) => {
    https.get(hljsCDN, res => {
      let data = [];
      res.on("data", buf => data.push(buf.toString("utf8")));
      res.on("end", () => resolve(data.join("")));
    }).on("error", reject);
  });

  script = "var self = {};\n" + script + "\nvar hljs = self.hljs;\n";
  await fs.writeFile(path.join(buildPath, hljsFilename), script, "utf8");
  return script;
}

async function highlighterStylemaps() {
  if (program.clean) {
    await clean();
  }

  let stylesMap;
  try {
    stylesMap = await fs.readFile(path.join(buildPath, stylesMapFilename), "utf8");
    if (stylesMap) {
      return stylesMap;
    }
  } catch (ex) {}

  const stylesPath = path.join(hljsPath, "styles");
  let styles = (await fs.readdir(stylesPath)).filter(file => {
    return file.endsWith("css") && !skipStyles.has(file);
  });
  
  stylesMap = {
    themeNames: []
  };
  for (let style of styles) {
    let contents = await fs.readFile(path.join(stylesPath, style), "utf8");
    let parsed;
    try {
      parsed = css.parse(contents);
    } catch (ex) {
      //
    }
    if (!parsed) {
      continue;
    }
    // Now add the rules we found to the map:
    let name = style.replace(/\.css$/, "");
    stylesMap[name] = {};
    stylesMap.themeNames.push(name);
    stylesMap[`$__funcPlaceholder__$${name}`] = name;
    for (let rule of parsed.stylesheet.rules) {
      if (rule.type != "rule") {
        continue;
      }
      let ruleBody = {};
      for (let decl of rule.declarations) {
        if (decl.type != "declaration") {
          continue;
        }
        ruleBody[decl.property] = decl.value;
      }
      // For each selector that this rule is applied to, add it to the map.
      for (let sel of rule.selectors) {
        stylesMap[name][sel] = ruleBody;
      }
    }
  }

  let strMap = "var hljsStyles = " + JSON.stringify(stylesMap, null, 2)
    .replace(/"\$__funcPlaceholder__\$.*".*"([a-z\d-]+).*"/g,
      (m, themeName) => {
        return `${camelize("change-" + themeName)}: function() { changeTheme("${themeName}"); }`
      }) +
    ";\n";
  await fs.writeFile(path.join(buildPath, stylesMapFilename), strMap, "utf8");
  return strMap;
}

async function highlighterScript() {
  let stylesMap = await highlighterStylemaps();
  let hljsScript = await downloadScript();

  await fs.writeFile(path.join(buildPath, hljsBundleFilename),
    stylesMap + hljsScript, "utf8");
}

async function highlighterUpdate() {
  await highlighterScript();
  await fs.copyFile(path.join(buildPath, hljsBundleFilename), path.join(rootPath, "documents", "highlighter.js"))
}

program
  .option("-c, --clean", "clean the build target before we start");

program
  .command("highlighter-stylemaps")
  .alias("stylemaps")
  .description("convert the highlighter.js CSS themes to JSON")
  .action(highlighterStylemaps);

program
  .command("highlighter-script")
  .alias("script")
  .description("create the full scripts that allows for highlighting codeblocks")
  .action(highlighterScript);

program
  .command("highlighter-update")
  .alias("update")
  .description("build and update the highlighter script(s) in the directories under source controlled")
  .action(highlighterUpdate);

// Error on unknown commands.
program.on("command:*", function () {
  console.error("Invalid command: %s\nSee --help for a list of available commands.", program.args.join(" "));
  process.exit(1);
});

program.parse(process.argv);
