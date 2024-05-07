## Automatic Email Link Opener

This Python script automates the process of checking your Outlook inbox for emails with specific criteria and opening the first relevant link it finds.

**Requirements:**

- Python 3.x
- `win32com` library (install using `pip install pywin32`)
- `webbrowser` library (included in the standard library)

**How to Use:**

1. **Install Dependencies:**

   ```bash
   pip install pywin32
   ```
2. **Run the Script:**
    ```bash
    python email_link_opener.py
    ```
**Configuration:**

The script prompts you for several configuration options:

- `DELTA MINS`: The number of minutes in the past to search for emails (default: 5).
- `SENDER`: The expected email sender address (default: "testemail@test.com").
- `SUBJECT`: The expected email subject line (default: "test subject").
- `URL KEYWORD`: A keyword to identify the relevant link within the email body (default: "https").
- `TYPE (B/E)`: Whether to find the Best Match (B) or Exact Match (E) for the URL keyword (default: "E").

**Functionality:**

- The script connects to your Outlook inbox using the win32com library.
- It searches for emails received within the specified delta timeframe that match the provided subject line and sender address.
- Depending on the match type (B or E), it identifies the first email containing a URL with or without the URL keyword.
- It extracts the identified link and opens it in your default web browser using the webbrowser library.

Notes:

- The script currently uses `webbrowser.open(url_to_go, new=0)`, which opens the link in the same browser window or tab if possible. Adjust this line if you prefer opening links in a new window. Uncomment the below lines and remove the open above line.
```
## customize to any browser as per needed

# edge_path = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe'
# webbrowser.register('edge', None, webbrowser.BackgroundBrowser(edge_path))
# webbrowser.get('edge').open(url_to_go)
```

- This script relies on the win32com library, which might require administrative privileges on certain systems.
- Error handling and logging could be further enhanced for better robustness.

Disclaimer:

This script is provided for educational purposes only. Use it at your own risk, and consider the security implications of automatically opening email links.