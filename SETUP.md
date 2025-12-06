# Excel AI Copilot - Setup Guide

## 🚀 Two Ways to Use This Add-in

### Option A: Development Mode (localhost)
For testing and development on your machine.

### Option B: Production Deployment (Recommended)
Deploy once, use anywhere - like the built-in Copilot!

---

## Option A: Development Mode

### 1. Get a Gemini API Key
1. Go to [Google AI Studio](https://aistudio.google.com/apikey)
2. Create a new API key
3. Copy the key

### 2. Install & Run
```bash
npm install
npm run start
```

### 3. Configure API Key in Excel
Click "AI Copilot" → Enter your API key when prompted

**Security Note:** API keys are stored with basic obfuscation in browser storage. For maximum security:
- Use the "Remove API Key" button in Settings when not actively using the add-in
- Consider re-entering your API key each session rather than storing it
- Never share your API key or workbooks that might contain stored keys

---

## Option B: Production Deployment (GitHub Pages - FREE)

### Step 1: Create GitHub Repository
1. Go to [github.com/new](https://github.com/new)
2. Create a new repo named `excel-ai-copilot`
3. Make it **Public** (required for free GitHub Pages)

### Step 2: Build & Push
```bash
# Build production files
npm run build

# Initialize git (if not already)
git init
git add .
git commit -m "Initial commit"

# Push to GitHub
git remote add origin https://github.com/YOUR_USERNAME/excel-ai-copilot.git
git branch -M main
git push -u origin main
```

### Step 3: Enable GitHub Pages
1. Go to your repo → Settings → Pages
2. Source: **Deploy from a branch**
3. Branch: **main** → **/dist** folder
4. Click Save
5. Wait 2-3 minutes for deployment

### Step 4: Update Manifest
1. Open `manifest.prod.xml`
2. Replace all `YOUR_GITHUB_USERNAME` with your GitHub username
3. Replace all `YOUR_REPO_NAME` with `excel-ai-copilot`
4. Save the file

### Step 5: Install the Add-in Permanently
**For Personal Use:**
1. Open Excel → Insert → Get Add-ins → My Add-ins
2. Click "Upload My Add-in"
3. Upload your `manifest.prod.xml`

**For Organization-wide:**
1. Go to Microsoft 365 Admin Center
2. Settings → Integrated Apps → Upload custom apps
3. Upload `manifest.prod.xml`

Now the add-in works in ANY workbook, on ANY device with your Microsoft account!

---

## Features

### Quick Actions
- **Analyze Data** - Get insights, patterns, and statistics
- **Create Formula** - Generate Excel formulas from natural language
- **Create Chart** - Get chart recommendations and create visualizations
- **Format Data** - Apply formatting and conditional formatting
- **Clean Data** - Find duplicates, fix inconsistencies
- **Summarize** - Create summaries with totals and averages

### Chat Interface
- Ask any question about your data
- Request specific formulas or calculations
- Get explanations of Excel functions
- Ask for data transformation help

### Apply Changes
When the AI suggests modifications, click "Apply Changes to Sheet" to execute them.

## Usage Tips

1. **Select data first** - The AI works best when you select the relevant data range
2. **Be specific** - "Sum column B" works better than "add up the numbers"
3. **Use quick actions** - They're optimized prompts for common tasks
4. **Review before applying** - Always check the AI's suggestion before applying

## Security

### API Key Storage
- API keys are stored with basic obfuscation (base64 encoding) in browser localStorage
- This is NOT encryption - it only prevents casual viewing
- For maximum security, use the "Remove API Key" button in Settings when done
- You may need to re-enter your API key after clearing browser data

### Removing Your API Key
1. Open Settings (gear icon)
2. Click "Remove API Key" button
3. Your key will be cleared from storage immediately

## Troubleshooting

### Add-in doesn't load
- Ensure you're running `npm run start`
- Check that port 3000 is available
- Try clearing the Office cache

### API errors
- Verify your Gemini API key is correct
- Check your API quota at Google AI Studio
- Ensure you have internet connectivity

### Changes not applying
- Make sure you have a cell/range selected
- Check the status bar for error messages
- Open the Diagnostics panel (document icon) to see detailed logs

### Debugging Issues
1. Click the Diagnostics button (document icon) in the header
2. Enable Debug Mode in Settings for more verbose logging
3. Check the logs for specific error messages and skipped actions

## Development

### Build for production
```bash
npm run build
```

### Validate manifest
```bash
npm run validate
```

## Deployment

To deploy for organization-wide use:
1. Build the production version
2. Host the files on a web server with HTTPS
3. Update manifest.xml URLs to point to your server
4. Deploy via Microsoft 365 Admin Center or SharePoint App Catalog
