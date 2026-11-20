# M K Fashion Vendor Payables

A dashboard for vendor payments to Greige suppliers and Process Houses. **The Excel workbook in OneDrive is the master**: the accountant keeps working in Excel, and the app reads the saved file automatically and shows it to the whole team. It works on phone, iPad and laptop, uses Google sign in with a personal PIN, and keeps a synced copy in Google Cloud Firestore.

## What is inside

| File | Purpose |
|---|---|
| `index.html` | The page shell: sign in, PIN lock, side panel |
| `css/styles.css` | Theme taken from the Excel workbook (maroon, salmon, teal) |
| `js/config.js` | **Your Firebase settings go here** (Step 2 below) |
| `js/calc.js` | All business rules: due dates, status, 12% greige interest, aging |
| `js/app.js` | All screens |
| `js/store.js` | Reads and writes data (Firebase, or demo mode in the browser) |
| `js/importer.js` | Reads the workbook and works out what changed since the last sync |
| `js/onedrive.js` | Microsoft sign in and OneDrive file access |
| `ms-auth.html` | Small page that completes the Microsoft sign in |
| `firestore.rules` | Security rules (Step 6 below) |

## Screens

1. **Vendor Payment Dashboard** with the Greige / Process House switch: KPI cards, status breakdown, aging schedule, vendor wise payables and charts.
2. **Process House** and **Greige**: bill lists with search, status filters and sorting. Tap a bill to see its instalments and the interest working. Bills are read only here: they are changed in Excel.
3. **Vendor Master List**: names and categories come from the Excel VENDOR MASTER sheet. Contact, GSTIN, PAN, address, bank and credit terms are entered in the app and kept there.
4. **Payment Planner**: pick a payment date, see what falls due, tick bills, and get the cash needed including greige interest. Export to CSV or print.
5. **Vendor Ledger**: running account statement for any vendor across all years, with opening balance by date range. Print or save as PDF for balance confirmation.
6. **Excel Sync** (admin and editor): OneDrive connection, workbook choice, last changes, points to check in the Excel, export and backup.
7. **Settings and Users** (admin only): interest rate, grace days, default terms, team members with roles, activity log.

## Business rules (same as the corrected Excel)

1. Due date = invoice date + credit terms.
2. Balance = invoice amount minus all payments, TDS and debit note.
3. Status: Paid on time, Paid after due date, Need urgent attention (open and past due, after 5 grace days for Process House, none for Greige), Partially paid, Pending.
4. Greige interest at 12% p.a. on a 365 day year: on each instalment paid after the due date, plus accrued on the open balance of an overdue bill. Process House bills carry no interest.

---

# Part A. Firebase and GitHub (already done)

Allow about 30 minutes. You need a Google account and a GitHub account.

## Step 1. Create the Firebase project

1. Open https://console.firebase.google.com and sign in with the Google account that will own the app.
2. Click **Create a project** (or **Add project**). Name it, for example, `mk-fashion-payables`.
3. Google Analytics is not needed. Turn it off and click **Create project**.
4. Stay on the free **Spark** plan. No card is needed.

## Step 2. Register the web app and copy the config

1. On the project home page click the **Web** icon `</>` (Add app).
2. Nickname: `payables-web`. Leave "Firebase Hosting" unticked. Click **Register app**.
3. Firebase shows a block of code containing `const firebaseConfig = { apiKey: ..., authDomain: ..., ... }`.
4. Open `js/config.js` in this project and replace the six placeholder values with yours. Save the file.

These values are not secret. Your data is protected by sign in and the security rules.

## Step 3. Turn on Google sign in

1. Left menu: **Build > Authentication > Get started**.
2. Tab **Sign-in method > Google > Enable**. Choose your email as the project support email. **Save**.

## Step 4. Create the database

1. Left menu: **Build > Firestore Database > Create database**.
2. Location: choose **asia-south1 (Mumbai)**. This cannot be changed later.
3. Choose **Start in production mode** and click **Create**.

## Step 5. Add yourself as the first admin

The app only lets in people listed in the `users` collection, so the first admin is added by hand.

1. In Firestore, tab **Data**, click **Start collection**.
2. Collection ID: `users`. Click **Next**.
3. Document ID: your Google email address in **lowercase**, for example `naman@gmail.com`.
4. Add three fields:

| Field | Type | Value |
|---|---|---|
| `role` | string | `admin` |
| `active` | boolean | `true` |
| `name` | string | your name |

5. Click **Save**. Everyone else can be added from inside the app later.

## Step 6. Publish the security rules

1. In Firestore, open the **Rules** tab.
2. Delete everything there, paste the full contents of `firestore.rules`, and click **Publish**.

## Step 7. Put the app on GitHub Pages

1. Create a repository on GitHub, for example `mk-payables`. A **public** repository is fine, because the code holds no business data. GitHub Pages on a private repository needs a paid GitHub plan.
2. Upload all files of this folder (keep the folders `css` and `js`). **Never upload your Excel file or a backup file to the repository.**
3. Repository **Settings > Pages**. Source: **Deploy from a branch**, branch `main`, folder `/ (root)`. **Save**.
4. After a minute the address appears, for example `https://yourname.github.io/mk-payables/`.

## Step 8. Authorise the GitHub address in Firebase

1. Firebase: **Authentication > Settings > Authorized domains > Add domain**.
2. Enter `yourname.github.io` (just the domain, no `https://` and no folder). **Add**.

## Step 9. First sign in

1. Open your GitHub Pages address. Click **Sign in with Google** and choose the admin account.
2. Create your 4 to 6 digit PIN.

## Step 10. Add the team

**Settings and Users > Add team member**: enter the person's Google email and pick a role.

| Role | Can do |
|---|---|
| Viewer | See everything, change nothing |
| Editor | Edit vendor details, connect OneDrive and sync |
| Admin | Everything, plus delete, settings, team, activity log |

They open the same address, sign in with that Google account and create their own PIN. If someone forgets their PIN, an admin clicks **Reset PIN**. To remove access, click **Disable** or **Remove**. This takes effect immediately.

## Recommended extra protection

1. **Restrict the API key to your site.** Open https://console.cloud.google.com, choose the same project, go to **APIs and Services > Credentials**, open the "Browser key", set **Application restrictions** to **Websites**, and add two entries: `https://yourname.github.io/*` and `https://your-project.firebaseapp.com/*` (the second one is needed for the Google sign in window; use your own project id from `authDomain`). Save.
2. **Weekly backup.** Admins can click **Download full backup (JSON)** on the Import and Export page. Keep the file somewhere safe and private.

## Demo mode

Until `js/config.js` holds real Firebase values, the app runs in demo mode: no sign in, data kept only in that browser. Add `?demo=1` to the address to try things safely at any time.

## Free plan limits

The Spark plan allows 50,000 reads and 20,000 writes per day. Each time the app opens it reads the current financial year's bills plus all open bills, so a team of around ten people is comfortably inside the limit. Choosing "All years" in the FY selector reads every bill, so use it only when needed.

---

# Part B. Register the app with Microsoft (one time, about 15 minutes)

This gives the app permission to read your workbook from your OneDrive. It only ever **reads** files.

1. On a laptop, open https://portal.azure.com and sign in with the **same Microsoft account that holds the Excel in OneDrive**.
2. Microsoft no longer lets a personal account register an app without a directory. If the portal says you have no directory or subscription, click the option to **start a free Azure account** and complete it. Microsoft asks for a phone number and a card only to verify your identity. Registering an app costs nothing.
3. In the search bar at the top type **App registrations**, open it, and click **New registration**.
4. Fill in:
   * Name: `MK Fashion Payables`
   * Supported account types: choose **Personal Microsoft accounts only**. If that option is not shown, choose the option that says **Any Entra ID tenant and personal Microsoft accounts**.
   * Redirect URI: platform **Single-page application (SPA)**, address `https://rishab2851.github.io/YOUR-REPO-NAME/ms-auth.html`
     (exactly your app address with `ms-auth.html` at the end. The Excel Sync page in the app shows the exact address with a Copy button.)
5. Click **Register**. On the overview page copy the **Application (client) ID**.
6. Left menu **API permissions > Add a permission > Microsoft Graph > Delegated permissions**, search `Files.Read`, tick it, click **Add permissions**. `User.Read` is already there. No admin consent is needed.
7. On GitHub open `js/config.js`, click the pencil, and paste the client ID into `msConfig.clientId`. Leave `authority: 'consumers'`. Commit.

# Part C. Update the files on GitHub

1. Upload these files to the repository, replacing the old ones (same folders):
   `index.html`, `ms-auth.html` (new), `css/styles.css`, `js/app.js`, `js/calc.js`, `js/importer.js`, `js/onedrive.js` (new), `js/store.js`, `js/ui.js`, `firestore.rules`, `README.md`.
2. **Do not upload `js/config.js` from the zip**, because yours already holds your Firebase values. Instead open your `js/config.js` on GitHub, click the pencil, and paste the block from the end of the new `config.js` (from `export const msConfig` to `SYNC_CHECK_SECONDS = 30;`) at the bottom. Add your Microsoft client ID there. Commit.
3. Firebase console: **Firestore Database > Rules**, paste the new `firestore.rules`, **Publish**. (One rule changed so editors can save the sync record.)
4. Wait a minute for GitHub Pages, then reload the app. On a phone, close and reopen the tab.

# Part D. Put the workbook in OneDrive and connect

1. Save `MK_Fashion_2026_Vendor_Payments.xlsx` (the version with the PAYMENTS sheet) in your OneDrive, for example in an `Accounts` folder. Your accountant opens it **from OneDrive** from now on.
2. In Excel, switch **AutoSave on** (top left). Changes then reach OneDrive within seconds.
3. In the app go to **Excel Sync > Connect OneDrive** and sign in with the Microsoft account. Then **Choose workbook**, search for part of the name, and pick the file.
4. The first sync loads everything. If the app still has bills or payments from the earlier one time import that are not in this workbook, it lists them and asks before removing anything.
5. Connect OneDrive on each device you use (your phone and your laptop). Other team members do not need to connect: they see the data as last synced, with the time shown at the top.

## How the sync behaves

* While the app is open on a connected device, it checks OneDrive every 30 seconds and downloads the file only when a newer version is saved.
* A short summary pops up on every open device listing what changed: new bills, payments, corrected amounts, vendors added or made inactive.
* Anything that would **remove** data (a bill deleted from Excel, a payment row deleted) is never removed silently. The app lists it and waits for an admin to confirm. "Not now" keeps it until the Excel changes again.
* Each bill remembers which workbook it came from. A new workbook next year can be connected without touching last year's bills.
* Vendors removed from the Excel VENDOR MASTER are marked inactive, with their details kept.
* Microsoft limits the sign in for web apps to 24 hours. After that the top bar shows **Reconnect OneDrive**: one tap, and it usually completes by itself.

## For the accountant

1. Open the workbook from OneDrive and keep **AutoSave on**.
2. Enter bills on GREIGE and PROCESS HOUSE as before.
3. Enter every payment, or each instalment, as a new row on the **PAYMENTS** sheet. AMT PAID and PAID ON on the bill sheets fill in by themselves (grey columns, do not type there).
4. A new vendor goes into **VENDOR MASTER** first, then appears in the dropdowns.
5. A red cell in a **CHECK** column means something to fix: duplicate bill, missing data, or a payment that does not match a bill.
