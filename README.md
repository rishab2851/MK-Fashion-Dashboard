# M K Fashion Vendor Payables

A web app for managing vendor payments to Greige suppliers and Process Houses. It works on phone, iPad and laptop, uses Google sign in with a personal PIN, and keeps data in Google Cloud Firestore so the whole team sees the same figures live.

## What is inside

| File | Purpose |
|---|---|
| `index.html` | The page shell: sign in, PIN lock, side panel |
| `css/styles.css` | Theme taken from the Excel workbook (maroon, salmon, teal) |
| `js/config.js` | **Your Firebase settings go here** (Step 2 below) |
| `js/calc.js` | All business rules: due dates, status, 12% greige interest, aging |
| `js/app.js` | All screens |
| `js/store.js` | Reads and writes data (Firebase, or demo mode in the browser) |
| `js/importer.js` | Excel import and export |
| `firestore.rules` | Security rules (Step 6 below) |

## Screens

1. **Vendor Payment Dashboard** with the Greige / Process House switch: KPI cards, status breakdown, aging schedule, vendor wise payables and charts.
2. **Process House** and **Greige**: bill lists with search, status filters and sorting. Tap a bill to record payments (instalments supported), edit it, or see the interest working.
3. **Vendor Master List**: every vendor with full details (contact, GSTIN, PAN, address, bank, credit terms), grouped as Greige Supplier, Process House, Transport or Other.
4. **Payment Planner**: pick a payment date, see what falls due, tick bills, and get the cash needed including greige interest. Export to CSV, print, or mark the batch as paid.
5. **Vendor Ledger**: running account statement for any vendor across all years, with opening balance by date range. Print or save as PDF for balance confirmation.
6. **Import and Export** (admin and editor): import the Excel workbook, export everything to Excel, JSON backup and restore.
7. **Settings and Users** (admin only): interest rate, grace days, default terms, team members with roles, activity log.

## Business rules (same as the corrected Excel)

1. Due date = invoice date + credit terms.
2. Balance = invoice amount minus all payments, TDS and debit note.
3. Status: Paid on time, Paid after due date, Need urgent attention (open and past due, after 5 grace days for Process House, none for Greige), Partially paid, Pending.
4. Greige interest at 12% p.a. on a 365 day year: on each instalment paid after the due date, plus accrued on the open balance of an overdue bill. Process House bills carry no interest.

---

# Setup guide

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

## Step 9. First sign in and data import

1. Open your GitHub Pages address. Click **Sign in with Google** and choose the admin account.
2. Create your 4 to 6 digit PIN.
3. Go to **Import and Export**, choose your Excel workbook, review the preview, and click **Import now**.
4. Check the Dashboard against Excel.

## Step 10. Add the team

**Settings and Users > Add team member**: enter the person's Google email and pick a role.

| Role | Can do |
|---|---|
| Viewer | See everything, change nothing |
| Editor | Add and edit bills, payments and vendors, import Excel |
| Admin | Everything, plus delete, settings, team, activity log |

They open the same address, sign in with that Google account and create their own PIN. If someone forgets their PIN, an admin clicks **Reset PIN**. To remove access, click **Disable** or **Remove**. This takes effect immediately.

## Recommended extra protection

1. **Restrict the API key to your site.** Open https://console.cloud.google.com, choose the same project, go to **APIs and Services > Credentials**, open the "Browser key", set **Application restrictions** to **Websites**, and add two entries: `https://yourname.github.io/*` and `https://your-project.firebaseapp.com/*` (the second one is needed for the Google sign in window; use your own project id from `authDomain`). Save.
2. **Weekly backup.** Admins can click **Download full backup (JSON)** on the Import and Export page. Keep the file somewhere safe and private.

## Using Excel alongside the app

Re-importing the same workbook updates bills instead of duplicating them. The Excel AMT PAID figure is kept as one payment entry marked "Excel" and is replaced on each import. Payments recorded in the app are never touched. Once the team records payments in the app, stop recording them in Excel, otherwise they are counted twice. The import preview warns you if it sees this. A file exported from the app cannot be imported back, for the same reason.

## Demo mode

Until `js/config.js` holds real Firebase values, the app runs in demo mode: no sign in, data kept only in that browser. Add `?demo=1` to the address to try things safely at any time.

## Free plan limits

The Spark plan allows 50,000 reads and 20,000 writes per day. Each time the app opens it reads the current financial year's bills plus all open bills, so a team of around ten people is comfortably inside the limit. Choosing "All years" in the FY selector reads every bill, so use it only when needed.
