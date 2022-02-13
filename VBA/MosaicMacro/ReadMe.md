# Read Me

This add-in helps you add color to your time at the office by transforming your Excel sheets into works of art.

## Setup Instructions

- Download the add-in files/clone the repository

- Enable the developer tab
    1. Press "File" in the top left
    2. Press "Options" in the bottom left
    3. Press "Customize Ribbons" on the left pane
    4. In the right window, check the line that says "Developer"
    5. Press "Ok"

- Add the add-in
    1. Navigate to the Developer tab in the Excel ribbon
    2. Press "Excel Add-ins"
    3. Press "Browse"
    4. Navigate to the location of the add-in and double click
    5. Ensure that the add-in shows up in the Excel add-ins window and that the line is checked

## Technical Notes

- Create Mosaic: creates blank canvas, resizes tiles, then colors the tiles
- Create Blank Canvas: clears all information in the worksheet, resizes all cells in the nxn grid, adds borders, and adds gray coloring to all cells outside the nxn grid
- Resize tiles: loops through every cell in the nxn grid, moving horizontally, and merges cells
- Color canvas: loops through every cell in the nxn grid, colors the top-left cell in the merged area, attempting to assign a unique color to each cell whenever possible
- Change settings: (in development) retrieves setting preferences via a userform and writes the preferences to a file in the same directory

- The bas files that contain the VBA code can be found in "Files\Bas Files"
- The XML files used to create the custom ribbon can be found in "Files\MyRibbon" And "Files\\_rels"