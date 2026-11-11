# 🌐 Vercel Deployment Procedure for MCO Maker

This guide walks you through deploying the **MCO Maker Licensing & Auto-Update Server** on **Vercel** ($0 Free Tier forever) and connecting it to your existing **Supabase** database.

---

## 📋 Overview of What We Are Deploying
- **Microservice:** Express.js Serverless Function (`license_server/`)
- **Database:** Supabase PostgreSQL (using multi-tenant tables: `mco_licenses`, `mco_license_devices`, `mco_app_updates`)
- **Endpoints:**
  - `POST /api/license/activate` - Hardware lock & subscription activation
  - `POST /api/license/verify` - Background heartbeat verification
  - `POST /api/admin/generate` - Generate new serial keys (protected by `ADMIN_SECRET`)
  - `GET /api/updates/latest` - Serve updates to Windows 10/11 and Windows 7 clients
  - `POST /api/admin/broadcast-update` - Broadcast new releases from `release.py`

---

## 🚀 Step 1: Initialize Database Tables in Supabase
1. Open your existing [Supabase Dashboard](https://supabase.com/dashboard).
2. Click on your project.
3. In the left sidebar, click **SQL Editor**.
4. Open a **New Query Tab** and name it `Tab 1: MCOMaker`.
5. Copy and paste the contents of `SUPABASE_QUERY.sql` (or `license_server/schema.sql`).
6. Click **Run** (Ctrl + Enter).
   - This creates `mco_licenses`, `mco_license_devices`, and `mco_app_updates` without touching any of your other existing tables (like Bill Scrapper)!
7. Open another **New Query Tab** and name it `Tab 2: Users MCOMaker`.
8. Paste the live user monitoring query from the bottom of `SUPABASE_QUERY.sql`. Keep this tab saved for inspecting active customer devices anytime!

---

## 🔑 Step 2: Retrieve Supabase API Credentials
1. In Supabase Dashboard, click **Project Settings** (gear icon at the bottom of the left sidebar).
2. Go to **API**.
3. Note down:
   - **Project URL:** `https://xxxxxxxxxxxxxxxx.supabase.co`
   - **Service Role Secret Key (`service_role`):** `eyJhbGciOi...`
   *(⚠️ Important: Do NOT use the `anon` key. The license server must use the `service_role` key to bypass RLS safely).*

---

## 🚢 Step 3: Deploy to Vercel

### Option A: Via GitHub (Recommended)
1. Push your `MCOMaker` repository to GitHub:
   ```bash
   git init
   git add .
   git commit -m "Initial commit for MCO Maker"
   git branch -M main
   git remote add origin https://github.com/<your-username>/MCOMaker.git
   git push -u origin main
   ```
2. Go to [Vercel Dashboard](https://vercel.com/dashboard) and click **Add New...** -> **Project**.
3. Select your `MCOMaker` GitHub repository and click **Import**.
4. In the configuration screen:
   - **Framework Preset:** Other
   - **Root Directory:** Click **Edit** and choose `license_server`
5. Expand **Environment Variables** and add:

| Key | Value | Description |
|---|---|---|
| `SUPABASE_URL` | `https://xxxx.supabase.co` | Your Supabase Project URL |
| `SUPABASE_SERVICE_ROLE_KEY` | `eyJ...` | Supabase Service Role Secret Key |
| `ADMIN_SECRET` | `YourChosenSecret2026` | Secret password for generating serial keys |
| `TABLE_LICENSES` | `mco_licenses` | Multi-tenancy licenses table |
| `TABLE_DEVICES` | `mco_license_devices` | Multi-tenancy devices table |
| `TABLE_UPDATES` | `mco_app_updates` | Multi-tenancy updates table |
| `KEY_PREFIX` | `MCO` | Serial key prefix (`MCO-XXXX-XXXX-XXXX`) |

6. Click **Deploy**.
7. In ~30 seconds, Vercel will give you a live URL, e.g.:
   `https://mco-maker.vercel.app` (or `https://mco-maker-yourname.vercel.app`).

### Option B: Via Vercel CLI
If you have Vercel CLI installed:
```bash
cd D:\Projects\MCOMaker\license_server
vercel
```
Follow the prompts, then set the environment variables via `vercel env add` or in the Vercel project web dashboard.

---

## 🧪 Step 4: Verify Deployment
Test your deployment by visiting the health endpoint in your browser or running in PowerShell:
```powershell
Invoke-RestMethod -Uri "https://mco-maker.vercel.app/health"
```
It should return:
```json
{
  "status": "ok",
  "app": "MCO Maker License Server",
  "tables": {
    "licenses": "mco_licenses",
    "devices": "mco_license_devices",
    "updates": "mco_app_updates"
  }
}
```

---

## 🛠 Step 5: Configure Admin Tool (`generate_key.py`)
1. Run the admin serial key generator tool:
   ```bash
   python generate_key.py
   ```
2. Click **⚙ Settings**:
   - **API Server URL:** Enter your Vercel URL (e.g. `https://mco-maker.vercel.app`)
   - **Admin Secret:** Enter your `ADMIN_SECRET`
3. Click **Save Settings**.
4. Select a subscription duration (e.g. 1 Year) and click **🔑 Generate Serial Key**!
   The serial key (e.g. `MCO-7K9W-2YXP-M4RF`) is automatically generated in Supabase and copied to your Windows clipboard!

---

## 🔄 Step 6: Link Desktop Client (`main.py`)
In `main.py`, the default license server URL is configured as:
```python
LICENSE_API_URL = os.environ.get("LICENSE_API_URL", "https://mco-maker.vercel.app").rstrip("/")
```
Once deployed, update this URL if your Vercel project name is different.
