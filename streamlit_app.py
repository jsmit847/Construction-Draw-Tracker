# Entry point for the main deployment, the app people open. Streamlit Community Cloud won't host two apps from
# the same repo, branch and file, and the original deployment (the approved Salesforce callback) runs
# draw_dashboard.py. Importing draw_dashboard, instead of running it, skips its OAuth relay block, so this app
# always finishes Salesforce logins itself.
from draw_dashboard import main

main()
