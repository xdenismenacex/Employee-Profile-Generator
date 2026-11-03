# Employee Profile Slides

A Python program that creates a PowerPoint presentation with a profile slide for each employee, showing their initials, name, job title, ID, department and start date.

## What it does

1. Reads employee details from a list.
2. Creates one widescreen slide per employee with a profile card design.
3. Adds a circle with the employee's initials and their details below it.
4. Saves everything as `profiles.pptx`, which you can open in PowerPoint, Google Slides or Keynote.

## How to run it

1. Open a new notebook at colab.research.google.com.
2. Paste the code into a cell.
3. Replace "ENTER VALUE" in the `people` list with real employee details.
4. Run the cell. The presentation downloads automatically.

The first line installs the `python-pptx` library, which the program needs to create PowerPoint files. To run it on your own computer instead, install the library with `pip install python-pptx` in your terminal, remove the `!pip` line and the last two lines (the `google.colab` part).

## Customising

- Add or remove employees in the `people` list. Each one gets their own slide.
- Change the colours at the top, for example `blue = RGBColor(0x2E, 0x7D, 0x32)` for green.

## What I learned

- Installing and using an external library
- Creating PowerPoint slides with shapes, text and colours
- Using nested loops
- Storing settings for several text lines in a list and looping through them
