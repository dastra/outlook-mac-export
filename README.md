This script allows you to back up Outlook for Mac (version 16) to your Mac's local hard drive.

It works using Applescript to open the Outlook client and back up.

You can either run it manually, or use launchd to run it on a schedule.

## Before running for the first time

Before running for the first time, you will need to 

1. Create a folder in your Documents folder called “Outlook Archives”
2. Download [the script](./outlook-export.scpt) and store it somewhere memorable

## Running the export script

To run the script:

```
osascript outlook-export.scpt
```
