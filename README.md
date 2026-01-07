# Railway Deployment Guide

## Step 1: Set Up Railway Account

1. Go to [railway.app](https://railway.app)
2. Click "Login" or "Start a New Project"
3. Sign up with GitHub (recommended for easy deployments)
4. Verify your email

**Free Tier:** $5 credit every month, no credit card required

---

## Step 2: Prepare Your Project

Your `railway/` directory structure should look like this:

```
railway/
├── app.py                  # Flask API (✅ created)
├── report_generators.py    # Report logic (✅ created)
├── requirements.txt        # Dependencies (✅ created)
├── Procfile               # Railway config (✅ created)
├── .gitignore            # Git ignore (✅ created)
├── README.md             # This file
├── scripts/              # Python report scripts
│   ├── qualys_report_automation_v3_2.py
│   └── executive_report_automation_v1_5.py
└── templates/            # Report templates
    ├── Template.xlsx
    └── Report-Template-with-Chart.docx
```

### Copy Files to Railway Directory

```bash
# Copy Python scripts
mkdir railway/scripts
copy "Detailed Report Files\qualys_report_automation_v3.2.py" railway\scripts\qualys_report_automation_v3_2.py
copy "Executive Report Files\executive_report_automation_v1.5_PRODUCTION.py" railway\scripts\executive_report_automation_v1_5.py

# Copy templates
mkdir railway/templates
copy "Detailed Report Files\Template.xlsx" railway\templates\Template.xlsx
copy "Executive Report Files\Report-Template-with-Chart.docx" railway\templates\Report-Template-with-Chart.docx
```

---

## Step 3: Deploy to Railway

### Option A: Using Railway CLI (Recommended)

1. **Install Railway CLI:**
   ```bash
   npm install -g @railway/cli
   ```

2. **Login:**
   ```bash
   railway login
   ```

3. **Initialize Project:**
   ```bash
   cd railway
   railway init
   ```
   - Enter a project name: "vm-monitoring-reports"
   - Choose "Empty Project"

4. **Deploy:**
   ```bash
   railway up
   ```
   
   This will:
   - Detect Python project
   - Install dependencies from requirements.txt
   - Start Gunicorn with your Flask app

5. **Get Your URL:**
   ```bash
   railway domain
   ```
   
   Copy the domain (e.g., `vm-monitoring-reports.up.railway.app`)

### Option B: Using GitHub (Alternative)

1. **Push to GitHub:**
   ```bash
   cd railway
   git init
   git add .
   git commit -m "Initial Railway deployment"
   git remote add origin https://github.com/your-username/vm-reports-api.git
   git push -u origin main
   ```

2. **Connect to Railway:**
   - Go to [railway.app/dashboard](https://railway.app/dashboard)
   - Click "New Project"
   - Click "Deploy from GitHub repo"
   - Select your repository
   - Railway will auto-deploy

3. **Generate Domain:**
   - Click on your service
   - Go to "Settings" tab
   - Click "Generate Domain"
   - Copy the URL

---

## Step 4: Test Railway Deployment

```bash
# Test health endpoint
curl https://your-app.railway.app/health

# Should return:
# {"status":"healthy","message":"Report generation service is running"}
```

---

## Step 5: Update Vercel Environment Variables

1. Go to [vercel.com/dashboard](https://vercel.com/dashboard)
2. Select your VM Monitoring project
3. Go to "Settings" → "Environment Variables"
4. Add new variable:
   - **Name:** `RAILWAY_API_URL`
   - **Value:** `https://your-app.railway.app`
   - **Environments:** Production, Preview, Development

5. Click "Save"

6. **Redeploy** your Vercel app to apply changes

---

## Step 6: Update Vercel API Routes

The API routes need to call Railway instead of local Python.

### Update `/api/reports/detailed/route.ts`

```typescript
const RAILWAY_URL = process.env.RAILWAY_API_URL

export async function POST(request: NextRequest) {
  const supabase = await createClient()
  const { data: { user } } = await supabase.auth.getUser()
  if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })

  const { scanId, businessUnit } = await request.json()

  // Fetch scan data
  const { data: scan } = await supabase
    .from('scans')
    .select('*')
    .eq('id', scanId)
    .single()

  const { data: vulnerabilities } = await supabase
    .from('vulnerabilities')
    .select(`
      *,
      assets!inner (*)
    `)
    .eq('scan_id', scanId)

  // Call Railway
  const response = await fetch(`${RAILWAY_URL}/api/reports/detailed`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      scan,
      vulnerabilities,
      businessUnit
    })
  })

  if (!response.ok) {
    return NextResponse.json({ error: 'Report generation failed' }, { status: 500 })
  }

  // Stream file to user
  const blob = await response.blob()
  return new NextResponse(blob, {
    headers: {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': `attachment; filename="detailed_report.xlsx"`
    }
  })
}
```

Similar changes for `/api/reports/executive/route.ts`

---

## Step 7: Test End-to-End

1. Start Next.js dev: `npm run dev`
2. Navigate to a scan: `http://localhost:3000/scans/[id]`
3. Click "Generate" on Detailed Report
4. Wait for Railway to process
5. Download the generated Excel file

---

## Monitoring & Debugging

### View Railway Logs

```bash
railway logs
```

Or in the Dashboard:
- railway.app/dashboard → Your Project → "Deployments" tab

### Common Issues

**Issue: "Module not found"**
- Solution: Check `requirements.txt` has all dependencies

**Issue: "Template not found"**
- Solution: Verify templates are in `railway/templates/`

**Issue: "Out of memory"**
- Solution: Upgrade to Railway Pro plan ($5/month base)

---

## Cost Optimization

### Free Tier Usage
- $5 credit = ~500 execution hours/month
- Each report takes ~30 seconds
- **Estimate:** ~6000 reports/month on free tier

### If You Exceed Free Tier
- Railway Pro: $5/month base + usage
- Vercel will cache some requests
- Consider implementing rate limiting

---

## Next Steps

After successful deployment:

1. ✅ Test with multiple scans
2. ✅ Verify report accuracy
3. ✅ Set up monitoring/alerts
4. ✅ Document for your team
5. ✅ Consider adding request queue for high volume

---

## Rollback Plan

If Railway doesn't work, you can fallback to local Python:

1. Revert Vercel API routes to use `pythonExecutor.ts`
2. Run Python service on your VM
3. Expose via HTTP endpoint

---

## Production Checklist

- [ ] Railway deployed and accessible
- [ ] Vercel environment variables set
- [ ] API routes updated
- [ ] End-to-end test successful
- [ ] Error handling tested
- [ ] Logs monitored for 24 hours
- [ ] Documentation updated
- [ ] Team training completed
