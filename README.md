# Print Station

A local or public printing website where users select a **room**, enter their
**name**, and upload files. An administrator can review and print all files in
a room.

## Features

- Upload one or more files with a room and name.
- Password-protected administration for managing rooms and reviewing,
  downloading, previewing, and deleting files.
- Print a room using either:
  1. **Print View (Browser Printing):** open the room's files in a browser and
     print them with `Ctrl+P`.
  2. **Direct Server Printing:** silently submit the room's files to a printer
     connected to the server computer.

## Requirements

- Windows and Python 3.9 or later.
- Microsoft Office (Word, Excel, and PowerPoint) to convert or print Office
  documents.
- Dependencies installed from `requirements.txt` (Flask, Waitress, and
  pywin32).

## Start the Site

Double-click the [startup batch file](启动打印站.bat). On the first run it
installs any missing Python packages, creates a unique administrator password,
starts the website, and prints the password in its console window.

Keep that console window open while the site is in use. To run it manually:

```powershell
cd C:\Users\1\Desktop\打印
pip install -r requirements.txt
python run.py
```

Open <http://127.0.0.1:5000> to upload files or
<http://127.0.0.1:5000/admin> to sign in as an administrator.

The administrator password is generated on first run and stored locally in
`.admin_password`. To change it, edit that file or set the `ADMIN_PASSWORD`
environment variable before starting the application. Do not share or commit the
password.

## Printing

### Browser Printing

Open a room in the admin area, select **Print View (Browser Printing)**, and
press `Ctrl+P`. Office documents are converted to PDF for preview.

### Direct Server Printing

Select a printer on the room page (or leave the system default selected), then
click **Print All Files in This Room**.

- Word, Excel, PowerPoint, TXT, and CSV files are printed or converted using
  Microsoft Office.
- Images (PNG, JPG, BMP, GIF, TIFF) are printed with GDI+, which needs no file
  association and forces A4 with the image fitted to the printable area. WebP
  images depend on GDI+ support in the installed Windows version.
- PDFs use the first available option:
  1. **SumatraPDF**, if `SumatraPDF.exe` is present in `tools\`.
  2. Adobe Reader, if installed.
  3. Edge silent printing as a fallback.

The room page remembers the selected printer when you click
**Remember as Default Printer**.

### Paper Size and Job Names

- Print jobs are configured for A4 where the file format and print backend
  allow it. Word and Excel use A4 settings, PDFs use SumatraPDF's A4 fit option,
  PowerPoint files are converted to PDF before printing, and images are drawn
  onto an A4 GDI+ page.
- Printer queue job names use `Room_Name` (for example, `301_Alex.docx`). If
  multiple files have the same room and name, a number is added to later jobs.
  The original file names remain visible on the website.
- Office and image jobs show `Room_Name.ext` in the printer queue. PDF jobs are
  rendered by SumatraPDF, which labels the job with the temporary file path, so
  the printer shows something like `...\ps_1a2b3c\301_Alex.pdf`.
- Each file is reported separately. A file is marked successful only when its
  print backend really accepted the job, for example when SumatraPDF exits with
  status 0 or Word finishes submitting the document. Failures include the reason,
  such as a missing printer name or an unavailable driver.
- Some printers and print servers delete jobs from the Windows queue as soon as
  they are transferred to the device, so a job that is waiting at the printer for
  a card swipe may no longer be listed on the computer.

### School Card-Release Printing

For a school printer that holds jobs until a teacher card is scanned:

1. Select the school printer on the room page.
2. Click **Print All Files in This Room**.
3. Go to the printer and swipe the teacher card to release the submitted jobs.

The current workstation was configured for a Canon imageRUNNER ADVANCE C5840.
When setting up another computer, add the school printer in Windows and select
it in the site.

## Public Access with cpolar

The configured public address is <https://print123.cpolar.cn>.

Start the site and tunnel by double-clicking the [startup batch file](启动打印站.bat).
The tunnel starts only when cpolar is installed and an authtoken is configured
in `%USERPROFILE%\.cpolar\cpolar.yml`; otherwise the site runs locally only.

The reserved name is read from [cpolar-subdomain.txt](cpolar-subdomain.txt).
Delete that file to let cpolar assign a random public URL instead.

Keep both console windows open, and keep the computer awake and connected to
the network while the site is in use.

To check the current public URL, run:

```powershell
python show_url.py
```

## Data and Backups

- Uploaded files: `uploads\`
- File metadata and room settings: `data.db`
- Preview cache: `uploads\.preview\`

Back up the application by copying its project folder while the service is
stopped.

## Git Version Management

This project is connected to the GitHub repository
`Andy123456789101112131415/Printer`.

Run `python autosave.py` to monitor changes and automatically commit and push
them after a 60-second quiet period. You can provide a different quiet period,
for example `python autosave.py 30`.

Uploaded files, the database, secret keys, and the admin password are excluded
from Git by `.gitignore`.

## Security Notes

- The admin area uses password-based sessions and CSRF protection.
- Each upload request is limited to 100 MB, and file extensions are restricted
  to an allowlist.
- Stored upload names are randomized, and room names are sanitized.
