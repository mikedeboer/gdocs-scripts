# gdocs-scripts

Collection of scripts that I use to automate things in Google Apps.


## How to use

The steps to follow are the same for each app (Docs or Sheets) and are manual,
unfortunately. However, they only consist of copying the files from this project
into Google's Script editor.

1. Open the app (Docs or Sheets) and select the menu 'Tools > Script editor',
2. Inside the newly opened Script editor, rename the project from 'Untitled project' to something you'd like,
3. Select 'File > New > Script file'
4. In the 'Create File' modal dialog, enter the filename of the file you'd like to copy from this repo (e.g. 'integration.gs')
5. Repeat 3 and 4 for all the files in the 'Docs' or Sheets' directory, depending on the app you're using.

For each app, the following features will become available:

## App: Docs.

 - At the end of the main menu, a new item 'Code formatter' will appear,
   - The first item in that menu is 'Selection as code', which when selected will transform the text you selected into pretty, highlightted code.
   - By default, the highlighter will attempt to detect which language the code is written in, but you can force a language by wrapping your code in a Github flavored Markdown block, using backticks.
   - The second item in that menu is 'Change theme', which hosts a submenu listing all available themes to choose from.
   - The Google Apps scripting API for menus is not rich enough to allow for form-like behavior inside the menus, so for now you'll have to believe & check that changing a theme 'just works'.

## App: Sheets

   - At the end of the main menu, a new item 'Manager Tools' will appear,
     - The first item in the menu is 'Generate Velocity Sheet', which when selected will generate a new sheet for the current date containing tables and graphs to illustrate the velocity of your team in past, present and future.
     - There's much to explain here, since a lot of data is presented and needs a couple of very specific steps to setup, which are better put in its own Readme.
