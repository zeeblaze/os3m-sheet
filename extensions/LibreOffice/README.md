# OS3M Sheet - LibreOffice Extension

This directory contains the source files for the LibreOffice extension component.

## File Structure

*   `Addons.xcu`: Configuration file for the LibreOffice UI. It registers the "OS3M Sheet" menu and links it to the Python script.
*   `Scripts/python/main.py`: The main Python logic for the extension. It handles the UNO interface, dialog creation, and communication with the backend server.

## Packaging

To create an installable `.oxt` extension file:
1.  Select the all contents of this directory (including `Addons.xcu` and the `Scripts` folder).
2.  Zip the selected files.
3.  Change the file extension from `.zip` to `.oxt`.

## Installation

To install the extension:
1.  Open LibreOffice.
2.  Go to **Tools** > **Extensions**.
3.  Click **Add**, select the `.oxt` file, and restart LibreOffice.