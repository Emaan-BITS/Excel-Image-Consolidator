# Excel Image Consolidator

A Python and Flask-based web application that automates the tedious process of inserting product images into Excel spreadsheets. The application reads product codes from your Excel file, finds the matching images from an uploaded folder, compresses them, and inserts them perfectly formatted into the spreadsheet.

## Live Demo
[https://MEmaanMI.pythonanywhere.com](https://MEmaanMI.pythonanywhere.com)

## Features
* **Automated Image Mapping:** Matches product codes in a specified Excel column to your uploaded image files.
* **Smart Image Compression:** Automatically resizes and converts images to optimized JPEGs, preventing your Excel files from becoming too large or slow to open.
* **Batch Processing:** Supports uploading entire directories of images and multiple Excel files simultaneously.
* **Real-Time Progress Tracker:** Displays a live progress bar so you know exactly what the application is doing during large tasks.
* **Automatic Formatting:** Adjusts Excel row heights and column widths dynamically to fit the inserted images neatly.
* **Missing Image Report:** Generates a downloadable text log detailing any product codes that did not have a matching image.

## Prerequisites
Ensure you have Python 3.7 or higher installed on your computer.

## How to Use

1. **Configure your settings:** * **Code Column:** Type the letter of the Excel column that contains your product IDs (e.g., A). 
   * **Image Size:** Select how large you want the images to appear inside the Excel file.

2. **Upload your files:** * **Step 1:** Upload the folder that contains all of your product images. 
   * **Step 2:** Upload your target Excel file(s) or an entire folder of Excel files.

3. **Consolidate:** Click **Consolidate Files**. A progress bar will appear to track the formatting and compression process. Once finished, a `.zip` file will automatically download to your computer containing your newly formatted Excel spreadsheets and the missing images report.
