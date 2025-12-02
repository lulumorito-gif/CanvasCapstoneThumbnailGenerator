# Capstone Thumbnail Automation Script 

This script automates the process of generating and uploading thumbnail images for Poster PDFs stored in SharePoint.

A floating “Make Thumbnails” button is injected into the page and allows you to process selected rows in bulk.

## 🧭 How to Use the Script

1. Open the SharePoint List
   
2. Navigate to the list or library containing Capstone items e.g.,2024 dataset.

3. Switch to Grid View
   
4. Ensure the list is in “Edit in grid view” mode before running the script.

5. Open the Browser Console
- Windows: `F12`
- Mac: `Cmd + Option + J`

6. Navigate to the Console tab.

7. Paste the Script.
  
Copy and paste the full script into the console and press Enter.

A floating blue button labeled “Make Thumbnails” will appear at the bottom-right corner.

8. Select the Rows to Process

Click to select one or more rows.
Each selected row must contain:
- A valid Poster link 
- SharePointPDFlink
- An empty or outdated PosterThumbnail value

9. Generate Thumbnails
Click the “Make Thumbnails” button or press Alt + Shift + T.
The script will:
- Double-click the PosterThumbnail cell
- Fetch a preview image from SharePoint
- Generate a 180×180 JPEG thumbnail
- Upload it automatically
- Continue to the next selected row

10. Remove the Floating Button
Right-click the button to remove it.

## ⚙️ What the Script Can Do

### ✅ Fully Automates Thumbnail Creation
- Generates and uploads thumbnails for each selected row.
- Uses SharePoint’s built-in upload dialogs and respects normal field-edit behavior.

### ✅ Works in Grid View
- Operates directly on inline editable cells in grid view.

### ✅ Handles SharePoint PDF Links
- Uses Microsoft Graph 
‘/shares/.../thumbnails/0/small/content‘ to fetch images.
- “Warms up” PDFs to encourage SharePoint to render thumbnails.

### ✅ Avoids Wrong UI States
- Detects and closes accidental “New Item” pop-up forms.
- Applies internal timing to help dialogs stabilize.

### ✅ Rich Console Logging
- Logs detailed progress \[Row\], \[Upload\], \[Warmup\], \[Retry\], \[ImgCheck\].

## 🚫 What the Script Cannot Handle

### ❌ Box.com Poster Links
- The current script only supports SharePoint URLs.
- Box URLs 
‘https://utexas.box.com/...‘ are skipped and logged as:
```
Row Poster is not SharePoint; skipping.
```
- To support Box links, you must use a Box Thumbnail Proxy Worker.

### ❌ Restricted or Broken SharePoint PDFs
- Password-protected or permission-restricted PDFs cannot be thumbnailed.

### ❌ Cross-Origin Files
- Files outside SharePoint 
Dropbox,GoogleDrive,Box,etc. are not supported.

### ❌ Unselected Rows
- Script only processes manually selected rows.
If no rows are selected, nothing happens.

## 🧩 Best Practices

- Make sure PDFs are accessible with your SharePoint permissions.
- Test with one row first before running on many.
- Keep the console open to watch logs and catch errors.
- Do not switch views or interact heavily with the UI while the script runs.
