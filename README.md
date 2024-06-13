# OutlookSignatureAddIn

Basic add-in for Outlook to show how to cache data in the browser session. Can be used on desktop and web versions of outlook.

Able to load signatures from Xink and store them in the local browser DB for later access.

Can only be access from the reading pane.

Domain Token must be entered and saved on first use.

Uses a custom API endpoint as a relay to Xink to avoid issues with CORS in local Dev, this applies the CORS policy to allow connection from localhost:3000
