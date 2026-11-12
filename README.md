A desktop application integrated with Outlook to prioritize emails, manage requests, and track daily meetings in one simple interface.

# BriefDesk – Outlook Desktop Productivity Application

BriefDesk is a Python-based desktop productivity application which integrates with Microsoft Outlook to organize emails, prioritize messages, categorize approval and vacation requests, and display daily calendar meetings through a simple, user-friendly interface.

## Output Screenshots

<img width="951" height="494" alt="Image" src="https://github.com/user-attachments/assets/2bb52f24-e878-47ad-8f6d-e3f485e26c53" />

<img width="958" height="498" alt="Image" src="https://github.com/user-attachments/assets/a5c8693b-225e-49b4-b283-ed8b6c8901a7" />

<img width="951" height="501" alt="Image" src="https://github.com/user-attachments/assets/84edeeaa-b808-4622-b9af-eb780b39c683" />

<img width="956" height="500" alt="Image" src="https://github.com/user-attachments/assets/fc796aec-0636-4670-8151-7fd938ca57a5" />

## Project Structure

- **mail-Brief-tool/**
  - `app.py` – Main Python application and Outlook integration
  - `index.html` – Application interface and JavaScript functionality
  - `style.css` – Interface styling and themes
  - `icon.ico` – Application system tray icon
  - `launch.bat` – Windows application launcher
  - `README.md` – Project documentation
    

## Technologies Used

- Python    –     Application logic and Outlook integration
- HTML      –     User interface structure
- CSS       –     Interface styling
- PyWebView –     Desktop window and Python–JavaScript integration
- PyWin32   –     Microsoft Outlook integration through COM
- Pillow    –     Icon image handling


## Features

- Email Management:         Retrieves emails received within the last 24 hours.
- Priority Classification:  Automatically categorizes emails into high, medium, and low priority based on keywords.
- Approval Requests:        Identifies emails related to approvals and authorization requests.
- Vacation Management:      Categorizes emails related to leave, vacations, and availability.
- Meeting Status:           Highlights ongoing, upcoming, and completed meetings (Based on priority).
- Quick Email Access:       Opens selected emails and calendar items directly in Outlook.
- Desktop Interface:        Uses HTML and CSS with PyWebView to provide a desktop interface.



## Prerequisites

- Windows operating system
- Python 3.9 or later
- Microsoft Outlook desktop application


## License

Specify an appropriate license before distributing or reusing this project.




