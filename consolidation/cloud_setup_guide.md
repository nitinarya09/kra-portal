# How to Host the KRA Compiler Online (For Free)

To take the compiler completely online so it doesn't require a local running PC, we will deploy the Python script to a free cloud hosting service named **Render** (render.com).

---

## Step 1: Create a GitHub Repository
1. Log in to your GitHub account (or create one at github.com).
2. Create a new repository named `kra-compiler`.
3. Set the repository visibility to **Private** (to protect your `credentials.json` and KRA Word template file).
4. Upload all files from your `d:\BUILDING and TESTING\KRA Automation\consolidation` folder to this repository:
   - `fetch_data.py`
   - `populate_template.py`
   - `appreciation_note.py`
   - `consolidate.py`
   - `cloud_server.py`
   - `requirements.txt`
   - `credentials.json`
   - `token.json` (your cached credentials)
   - Also upload the template file: `Sectionwise Blank KRA Report for June 2026 end Quarter.docx` into the repository.

---

## Step 2: Set up a Free Web Service on Render
1. Create a free account at [Render](https://render.com/).
2. Click the **+ New** button > select **Web Service**.
3. Connect your GitHub account and select your `kra-compiler` repository.
4. Set the following settings:
   - **Name**: `kra-compiler-service`
   - **Environment**: `Python`
   - **Region**: Select the closest region (e.g., Singapore or Oregon).
   - **Branch**: `main`
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `python cloud_server.py`
   - **Instance Type**: **Free** ($0/month)
5. Click **Create Web Service**.

Render will build the project and deploy it. Once complete, it will give you a public URL like:
`https://kra-compiler-service.onrender.com`

---

## Step 3: Connect your Web Portal to the Cloud

Once your Render service is live, update the Web Portal to fetch from Render instead of localhost.

In `d:\BUILDING and TESTING\KRA Automation\frontend\src\components\Dashboard.jsx` (near line 172):

```javascript
// Replace this:
const response = await fetch(`http://localhost:5005/compile?fy=${selectedFY}&quarter=${selectedQuarter.split(' ')[0]}`);

// With your live Render URL:
const response = await fetch(`https://kra-compiler-service.onrender.com/compile?fy=${selectedFY}&quarter=${selectedQuarter.split(' ')[0]}`);
```

Rebuild the React project (`npm run build`) and host the static files on a free service like GitHub Pages.

---

## How it works after deployment:
1. Staff go to the web portal link on their phone/PC, fill in data, and submit.
2. The Admin logs in, selects the FY/Quarter on the dashboard, and clicks **Compile Master KRA Report (.docx)**.
3. The browser calls the Render cloud server.
4. Render fetches the data from your Google Sheet, merges it into the KRA template, and **directly triggers a file download of the completed KRA `.docx` file in the Admin's browser!**
