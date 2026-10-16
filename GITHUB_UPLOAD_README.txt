RANKPILOT — COMPLETE SOURCE PACKAGE v09

1. Download and extract this ZIP.
2. Open the extracted folder and upload the CONTENTS (files/folders at the root) to https://github.com/harshkumarcsoft/rankpilot-api. Do not upload the ZIP file itself.
3. Replace existing files when GitHub asks. Commit all changes to the main branch.
4. Render: wait for the rankpilot-api service to redeploy and become Live. Test https://rankpilot-api-ovfh.onrender.com/api/health.
5. Vercel: redeploy the frontend after GitHub updates. Ensure the project environment variable VITE_API_BASE_URL is set to https://rankpilot-api-ovfh.onrender.com/api for Production. The frontend also has this Render URL as a production fallback.
6. If project creation fails, the new UI error should include the actual API URL and HTTP status instead of the old generic JSON parse error.

Do not upload local databases, client reports, .venv, node_modules, or secrets.

Important: the API prototype currently has no authentication and local report/database storage may be ephemeral on Render. Do not use it for confidential client data until authentication and persistent storage are configured.
