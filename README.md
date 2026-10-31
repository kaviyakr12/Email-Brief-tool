A desktop application integrated with Outlook to prioritize emails, manage requests, and track daily meetings in one simple interface.

# BriefDesk – Outlook Desktop Productivity Application

BriefDesk is a Python-based desktop productivity application that integrates with Microsoft Outlook to organize emails, prioritize messages, categorize approval and vacation requests, and display daily calendar meetings through a simple, user-friendly interface.

## Features

- Email Management:         Retrieves emails received within the last 24 hours.
- Priority Classification:  Automatically categorizes emails into high, medium, and low priority based on keywords.
- Approval Requests:        Identifies emails related to approvals and authorization requests.
- Vacation Management:      Categorizes emails related to leave, vacations, and availability.
- Meeting Overview:         Displays daily calendar events, including meeting times, durations, organizers, and locations.
- Meeting Status:           Highlights ongoing, upcoming, and completed meetings.
- Quick Email Access:       Opens selected emails and calendar items directly in Outlook.
- Desktop Interface:        Uses HTML and CSS with PyWebView to provide a desktop interface.

## Technologies Used

- Python –     Application logic and Outlook integration
- HTML –       User interface structure
- CSS –        Interface styling
- PyWebView –  Desktop window and Python–JavaScript integration
- PyWin32 –    Microsoft Outlook integration through COM
- Pystray –    System tray functionality
- Pillow – Icon image handling


## Project Structure

Email-Brief-tool-/
│
├── app.py          # Main Python application and Outlook integration
├── index.html      # Application interface and JavaScript functionality
├── style.css       # Interface styling and themes
├── icon.ico        # Application system tray icon
├── launch.bat      # Windows application launcher
└── README.md       # Project documentation




## Prerequisites

- Windows operating system
- Python 3.9 or later
- Microsoft Outlook desktop application



## How It Works

1. The application connects to the Outlook desktop application.
2. It retrieves emails from the previous 24 hours and calendar meetings scheduled for the current day.
3. Emails are classified into general emails, approval requests, and vacation-related messages.
4. Priority levels are assigned based on predefined keywords and message metadata.
5. Meeting information is displayed with relevant timing and status indicators.
6. Users can open individual Outlook items and refresh the displayed information.

## Privacy and Security

BriefDesk is designed to retrieve and process Outlook information locally through the Outlook desktop application's COM interface.

The current Python implementation does not explicitly upload retrieved email or calendar information to an external server. However, the application's overall network behavior depends on the interface code, dependencies, and execution environment.

storing passwords, access tokens, or other sensitive information in project files should be avoided.

## Future Enhancements

- Search and filter emails by sender, subject, and priority.
- Add configurable priority rules.
- Improve email categorization accuracy.
- Introduce customizable meeting reminders.
- Enhance the interface with additional productivity features.

## License

Specify an appropriate license before distributing or reusing this project.


## Output Screenshots

<img width="951" height="494" alt="Image" src="https://github.com/user-attachments/assets/2bb52f24-e878-47ad-8f6d-e3f485e26c53" />

<img width="958" height="498" alt="Image" src="https://github.com/user-attachments/assets/a5c8693b-225e-49b4-b283-ed8b6c8901a7" />


