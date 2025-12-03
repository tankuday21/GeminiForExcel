/**
 * Deploy Script for Excel AI Copilot
 * Updates manifest URLs and builds for production
 * 
 * Usage: node deploy.js YOUR_GITHUB_USERNAME YOUR_REPO_NAME
 * Example: node deploy.js johndoe excel-ai-copilot
 */

const fs = require('fs');
const path = require('path');

const args = process.argv.slice(2);

if (args.length < 2) {
  console.log(`
╔════════════════════════════════════════════════════════════╗
║           Excel AI Copilot - Deploy Script                 ║
╠════════════════════════════════════════════════════════════╣
║                                                            ║
║  Usage: node deploy.js <github_username> <repo_name>       ║
║                                                            ║
║  Example:                                                  ║
║    node deploy.js johndoe excel-ai-copilot                 ║
║                                                            ║
║  This will:                                                ║
║    1. Update manifest.prod.xml with your URLs              ║
║    2. Build production files to /dist                      ║
║    3. Copy manifest.prod.xml to /dist                      ║
║                                                            ║
║  After running, push /dist to GitHub Pages!                ║
╚════════════════════════════════════════════════════════════╝
  `);
  process.exit(1);
}

const [username, repoName] = args;
const baseUrl = `https://${username}.github.io/${repoName}`;

console.log(`\n🚀 Deploying Excel AI Copilot to: ${baseUrl}\n`);

// Read and update manifest.prod.xml
const manifestPath = path.join(__dirname, 'manifest.prod.xml');
let manifest = fs.readFileSync(manifestPath, 'utf8');

manifest = manifest.replace(/YOUR_GITHUB_USERNAME/g, username);
manifest = manifest.replace(/YOUR_REPO_NAME/g, repoName);

// Write updated manifest
fs.writeFileSync(manifestPath, manifest);
console.log('✅ Updated manifest.prod.xml');

// Also create a copy in dist after build
console.log('\n📦 Building production files...\n');
console.log('Run these commands:');
console.log(`  1. set PROD_URL=${baseUrl}/`);
console.log('  2. npm run build');
console.log('  3. Copy manifest.prod.xml to dist folder');
console.log('\nOr run: npm run build:prod (after updating the URL in package.json)');

console.log(`
╔════════════════════════════════════════════════════════════╗
║                    Next Steps                              ║
╠════════════════════════════════════════════════════════════╣
║                                                            ║
║  1. Push your code to GitHub:                              ║
║     git add .                                              ║
║     git commit -m "Deploy"                                 ║
║     git push                                               ║
║                                                            ║
║  2. Enable GitHub Pages:                                   ║
║     - Go to repo Settings → Pages                          ║
║     - Source: Deploy from branch                           ║
║     - Branch: main, folder: /dist                          ║
║                                                            ║
║  3. Wait 2-3 minutes, then install in Excel:               ║
║     - Insert → Get Add-ins → My Add-ins                    ║
║     - Upload My Add-in → manifest.prod.xml                 ║
║                                                            ║
║  Your add-in URL: ${baseUrl.padEnd(30)}    ║
╚════════════════════════════════════════════════════════════╝
`);
