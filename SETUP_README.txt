ELECTORAL OPERATIONS STAFF LOGIN PACKAGE

Included:
- Code_updated.gs
- StaffLogin.html
- FieldDashboard.html
- RegionalDashboard.html
- NationalDashboard.html
- DashboardSessionGuard.html

QUICK SETUP
1. Back up your existing Code.gs first.
2. Copy Code_updated.gs contents into Code.gs in the existing Apps Script project.
3. Create Apps Script HTML files named exactly StaffLogin, FieldDashboard, RegionalDashboard, NationalDashboard and paste in the matching HTML file contents. Do not include ".html" in the Apps Script file names.
4. In Apps Script Project Settings > Script properties, configure STAFF_ACCOUNTS_JSON with each authorized demonstration account's role, passwordHash, question, and answerHash.
5. Generate hashes using hashSecretForSetup in Code.gs. Set TEMP_SECRET_TO_HASH to the secret; for a security answer also set TEMP_SECRET_KIND=answer; run the helper and copy the hash from the execution log. The helper removes the temporary properties after use.
6. Deploy > New deployment > Web app. Choose the appropriate execute-as and access settings for your organization, then deploy and open the /exec URL.
7. Test only with demonstration accounts first.

SECURITY NOTE
This is an independent prototype, not an official IEBC service. The login flow does not automatically secure every existing backend endpoint. Before any production use with sensitive data, add server-side session and role checks to every protected data read/write action, and conduct a security review. Security questions are weaker than strong MFA.
