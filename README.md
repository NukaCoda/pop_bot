# pop_bot
This program automatically flips through Pop Now cases and attempts to buy a box when available.

DO NOT MOVE .EXE FILE FROM FOLDER. Instead, make a shortcut.

Presets are unique for each computer, it is not recommended to share presets.

This program will control the user's cursor while it is running.

Long breakdown of program:
  This program prompts the user to enter data for where the boxes are located on the screen.
  Where the bottom of the case(the box that the smaller boxes are in) is located on the screen.
  Where the rough estimate of the error banner is located, this is what pops up when a box is
  clicked but is not actually available.
  Where the 'Next Case' button is, the sidways arrow to the right of the case of boxes.
  Where the bottom portion of the 'Next Case' button is, this is where the program will click.
  Where the serial number of the case is.
  Where the 'Pick One To Shake' button is, this is where the program will click when a box is available.
  And what color to look for to determine is a case is available(this should be a unique color that is
  only on an available case).

  The program will then run a few checks and click the next button automatically when the case is in
  position, the next button is loaded, and if there is no box available.

  Once a box is available it will click 'Pick One' and will then check to see if the error banner message appears.
  If there is an error it will wait a few seconds and then resume clicking 'Next' until another box is available.

  If the webpage times out, the screen changes, if that happens for more than 10 seconds, the program will 
  attempt to refresh the page. If the screen and the name of the browser window changes(potential error message from the website)
  the program will attempt to go back a page, wait, and then resume searching.

  {ESCAPE} is used to kill and close the program.
