# CanvasCapstoneThumbnailGenerator
2025 script to help automate processing numerous rows of capstone projects' posters into individual thumbnails for iSchool Capstone management.

## Usage guide 
Here’s a concise usage guide and feature summary for the current SharePoint thumbnail automation script.

🧭 How to Use the Script
Open the SharePoint List


Go to the list or library that contains your Capstone items (e.g., 2024 dataset).


Switch to Grid View


Ensure the list is in “Edit in grid view” mode before running the script.


Open the Browser Console


Press F12 (Windows) or Cmd + Option + J (Mac) to open Developer Tools → Console tab.


Paste the Script


Copy and paste the full script into the console, then press Enter.


A floating blue button labeled “Make Thumbnails” will appear at the bottom-right of the page.


Select the Rows to Process


Click to select one or more rows (they’ll be highlighted).


Each row must have:


A valid Poster link (SharePoint file link to a PDF)


Empty or outdated PosterThumbnail field


Generate Thumbnails


Click the “Make Thumbnails” button or press the shortcut Alt + Shift + T.


The script will:


Double-click the PosterThumbnail cell


Fetch a preview image from SharePoint


Create a 180×180 JPEG thumbnail


Upload it automatically


Move to the next selected row


Right-Click the Button to remove it when done.



⚙️ What the Script Can Do
✅ Fully Automates Thumbnail Creation
Generates and uploads thumbnails for each selected row in bulk.


Respects SharePoint’s built-in upload dialogs and confirmation steps.


✅ Works in Grid View
Detects and interacts directly with inline editable cells in grid mode.


✅ Handles SharePoint PDF Links
Extracts thumbnails via Microsoft Graph (/shares/.../thumbnails/0/small/content).


Includes “warm-up” logic to ensure PDFs render before fetching.


✅ Logs Everything
Outputs detailed [Row], [Upload], [Warmup], [Retry], and [ImgCheck] messages in the console.



🚫 What It Cannot Handle
❌ Box.com Poster Links
The script only supports SharePoint links (https://...sharepoint.com/...).


Box URLs will be skipped with a [Row] Poster is not SharePoint; skipping. message.
❌ Password-Protected or Inaccessible PDFs
If SharePoint permissions or file links are broken, the script cannot fetch thumbnails.


❌ Offline or Cross-Origin PDFs
It cannot read from outside SharePoint (e.g., Dropbox, Google Drive, etc.).


❌ Batch Mode Without Selection
It processes only manually selected rows.
 No selection = no action.



🧩 Best Practices
Test with one row before batch-running.


Keep the console visible to monitor progress.


Avoid editing or switching views while it’s running.

