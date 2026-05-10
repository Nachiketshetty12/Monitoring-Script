Omada AP Monitor
A Python script that polls a TP-Link Omada wireless controller, checks the status of all Access Points via ping, generates a styled Excel report, and sends an HTML email summary via SMTP.

Features
Authenticates with the Omada controller REST API
Fetches all APs across paginated results
Ping-checks each AP and marks it UP / DOWN
Maps MAC addresses to human-readable locations (via local JSON file)
Formats uptime from raw seconds into a readable string
Produces a colour-coded .xlsx report (green = online, red = offline)
Sends an Outlook-safe HTML email with a summary table and stats cards
Requirements
Python 3.8+
A reachable Omada Software Controller (tested on v5/v6)
Install dependencies:

pip install requests openpyxl ping3 python-dotenv
Setup
1. Clone the repo
git clone https://github.com/your-org/omada-ap-monitor.git
cd omada-ap-monitor
2. Configure environment variables
cp .env.example .env
Edit .env and fill in your values:

Variable	Description
OMADA_URL	Full URL of your controller, e.g. https://192.168.1.1:8043
OMADA_USER	Controller login username
OMADA_PASS	Controller login password
OMADA_SITE	Site name as shown in the controller UI
COMPANY	Label used in the report header
ZOHO_SMTP	SMTP host (default: smtp.zoho.com)
ZOHO_PORT	SMTP port (default: 587)
ZOHO_USER	SMTP sender address
ZOHO_PASS	SMTP password or app password
MAIL_TO	Comma-separated list of recipient addresses
MAC_INFO_PATH	Path to your mac_info.json file (default: mac_info.json)
3. Create your MAC → location map
cp mac_info.example.json mac_info.json
Edit mac_info.json with your actual AP MAC addresses, IPs, and location labels:

{
  "aa-bb-cc-dd-ee-01": {"ip": "192.168.1.10", "location": "Floor 1 - Room A"},
  "aa-bb-cc-dd-ee-02": {"ip": "192.168.1.11", "location": "Floor 2 - Room B"}
}
Note: mac_info.json and .env are listed in .gitignore and will never be committed.

Usage
Load your .env and run the script:

# With python-dotenv auto-loading (add to top of script if needed)
python ap_monitor.py

# Or export variables manually
export $(grep -v '^#' .env | xargs) && python ap_monitor.py
The script will:

Log in to the Omada controller
Resolve the configured site
Fetch and ping all APs
Build and save the Excel report
Send the HTML email with the report attached
Scheduling
To run automatically, add a cron job:

# Run every day at 8:00 AM
0 8 * * * cd /path/to/omada-ap-monitor && python ap_monitor.py >> logs/ap_monitor.log 2>&1
Project Structure
omada-ap-monitor/
├── ap_monitor.py          # Main script
├── mac_info.example.json  # Example MAC → location map (commit this)
├── .env.example           # Example environment config (commit this)
├── .gitignore
└── README.md
Security Notes
Never commit .env or mac_info.json — both are in .gitignore
Use an app password for SMTP rather than your account password where possible
The script disables SSL verification (verify=False) for the controller connection — consider adding your controller's self-signed cert to the trust store for production use
License
MIT
