# BFZ1-Google-Docs-Code
This guide will set up automation that synchronises shpreadsheet data onto a manual map

# Requirements, Assumptions and Considerations
- You will need the Google Slides and Google Sheets document IDs.
  - The code assumes you are working with two columns in a <ins>single spreadsheet tab</ins> and maps only to <ins>shapes</ins> in a <ins>single</ins> Google Slides document.
- The account on which the code is run from MUST have permissions to <ins>edit</ins> BOTH documents
- No element of this code, or usage tests, have been optimised for mobile use. PC is recommended.

# How it Works
Assuming no changes to the code are made and set up correctly, the code is run every minute, or, when the sync button is clicked in Google Sheets. It iterates all shapes in the Google Slides, if the name of a shape in the AltText field matches exactly with a field in Column A (`nameShape`), it will apply the colour of Column H (`val`) if defined. Because row 1 is a header row, A1 and H1 are ignored.

# Set Up
This section will set up and link the code and documents. A test example will also be provided which will be run using the exact code as presented.
## 1. Spreadsheet
You can start off with a blank Spreadsheet. Copy the Spreadsheet ID, which is found in the URL between the `d/` and `/edit` portions. Click on "Extentions" and then "App Scripts". Take note of the name of the Tab in the spreadsheet you wish to work with.

### For Our Example
Insert `Random Name` into cell A2 and `Purple Battle` into H2. Rename the tab `Station Claims`.

## 2. Shapes in Google Slides
Open a Google Slides document. Copy the Slides ID, found in the URL similar to the spreadsheet.

### For Our Example
Insert a shape. Right Click on the shape and select `Alt Text`. You can use the keyboard shortcut `Crtl+Alt+Y` if compatible. In the Advanced Options section, type `Random Name` into the field - this gives the shape a title. Close the formatting window and deselect the shape.

## 3. Google Scripts / App Scripts
Security Note: You may encounter at some stage in this step that a popup is blocked, or requires logging into the appropriate Google Account. It is safe to authorise this as it does not impact anything other than the two documents already define.

If a code file hasn't already been opened automatically, create one. Replace anything with the complete code. Those with code literacy will immediately be able to identify which values to customise - they are all commented.

In lines 15, 16 and 19, replace the names of the appropriate fields. Lines 55 to 65 define the colours to mapfor `val`. There is also a default colour if there exists a matching `shapeName` but no `val`. Colours use a hex code without the transparency value - i.e. `#rrggbb`

Save the code.

### For Our Example
We won't need to do anything, just make sure the IDs and Tab name matches.

## 4. Triggers
Warning: Depending on what you choose, the code may start to run immediately.

We now need to set up a trigger. In the left menu, Click on the alarm clock - it should bring you to the Triggers section.
In the bottom right, click Add Trigger.
For "Choose which function to run" select `updateSlidesfromSheet`
For "Choose which deployment should run" select `Head`
For "Select event source" and "Select event type" you have options:
* Time-driven event source allows you to run the code at defined times/intervals
* From spreadsheet event source allows you to run the code when the respective event happens on the spreadsheet.

Failure notifications can be chosen freely - immediately is recommended.

Click Save. Refresh the webpages. A button will appear on Google Sheets allowing you to manually sync the data, in addition to the settings of the trigger.

### For Our Example
Select Time-driven for the event source. Select "Minutes Timer" and then "Every Minute". Our code will now run every minute.
