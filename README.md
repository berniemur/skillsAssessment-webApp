**Note:** The link will generate an error when you enter it in chrome if you are signed into multiple chrome/gmail accounts. The easiest workaround it using a Chrome Incognito Tab or Safari Private Tab


**1. Create the Google Sheet**\
Sign in with the account that owns your Google Site and create a new Google Sheet, for example "Draft Scores 2026". Go to Extensions > Apps Script. Replace the starter code with Code.gs. Click + > HTML, name the file exactly Index, and paste in Index.html. Change ACCESS_CODE at the top of Code.gs to your own code, then save.\
\
**2. Set up the tabs**\
Reload the sheet. Choose the new menu Draft Tools > Set up sheets and approve the permission prompt. Paste your roster into the Roster tab (Player #, First, Last, Division) and your assessors' names into the Assessors tab.\
\
**3. Deploy as a web app**\
In Apps Script, go to Deploy > New deployment > Web app. Set Execute as to Me and Who has access to Anyone. The access code keeps outsiders out, and assessors won't need Google accounts or access to the sheet. Copy the URL that ends in /exec.\
\
**4. Embed it in your Google Site**\
In your Site editor, choose Insert > Embed > By URL, paste the /exec link, and insert it. Drag the embed box tall, ideally to full page height. Then publish the Site. You can also text assessors the /exec link directly, which gives a full-screen view on phones.\
\
**5. Score on assessment day**\
Each assessor enters the code once and taps their name. Then they tap 1–5 next to each player, filtering by division or searching by number. Every tap saves right away. If the gym signal drops, scores queue on the phone and send automatically when it comes back.\
\
**6. Pull results for the draft**\
Choose Draft Tools > Build / refresh Summary. You'll get one row per kid with every assessor's score, the average, and the spread. It's sorted by division and then by average, and flags kids who are missing scores or got very different scores. Use File > Download > Microsoft Excel if you want an .xlsx.
