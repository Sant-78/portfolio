# Portfolio GitHub Pages deploy script
$gitBin = "$env:LOCALAPPDATA\MinGit\cmd"
$ghBin = "$env:LOCALAPPDATA\GitHubCLI\bin"
$env:Path = "$gitBin;$ghBin;" + $env:Path

Set-Location "$PSScriptRoot"

Write-Host "`n=== GitHub Pages Deploy ===" -ForegroundColor Cyan

# Fix common Windows SSL push error (certificate revocation check)
git config --global http.schannelCheckRevoke false 2>$null
gh auth setup-git 2>$null

# Check GitHub login
$auth = gh auth status 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "GitHub login required. Browser will open..." -ForegroundColor Yellow
    gh auth login -h github.com -p https -w
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Login failed. Run this script again after authorizing GitHub." -ForegroundColor Red
        exit 1
    }
}

Write-Host "Logged in to GitHub." -ForegroundColor Green

# Ensure commit exists
if (-not (Test-Path ".git")) {
    git init
    git branch -M main
}
git add index.html css/ js/ .gitignore README.md deploy.ps1
if (Test-Path "resume.pdf") { git add resume.pdf }
git diff --cached --quiet
if ($LASTEXITCODE -ne 0) {
    git commit -m "Update portfolio site"
}

# Create repo if missing, then push
$repoExists = gh repo view Sant-78/portfolio 2>$null
if (-not $repoExists) {
    Write-Host "Creating repository Sant-78/portfolio..." -ForegroundColor Cyan
    gh repo create portfolio --public --source=. --remote=origin --push --description "Personal portfolio - Santosh Kumar Pal"
} else {
    Write-Host "Pushing to Sant-78/portfolio..." -ForegroundColor Cyan
    git remote remove origin 2>$null
    git remote add origin https://github.com/Sant-78/portfolio.git
    git -c http.sslVerify=false push -u origin main
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Push failed. Check your internet connection and run deploy.ps1 again." -ForegroundColor Red
        exit 1
    }
}

if ($LASTEXITCODE -ne 0) {
    Write-Host "Repository setup failed." -ForegroundColor Red
    exit 1
}

# Enable GitHub Pages
Write-Host "Enabling GitHub Pages..." -ForegroundColor Cyan
gh api repos/Sant-78/portfolio/pages -X POST -f "build_type=legacy" -f "source[branch]=main" -f "source[path]=/" 2>$null
if ($LASTEXITCODE -ne 0) {
    gh api repos/Sant-78/portfolio/pages -X PUT -f "build_type=legacy" -f "source[branch]=main" -f "source[path]=/" 2>$null
}

$pageStatus = gh api repos/Sant-78/portfolio/pages --jq '.status' 2>$null
if ($pageStatus) {
    Write-Host "`nDone! Your portfolio will be live in 1-3 minutes at:" -ForegroundColor Green
    Write-Host "https://sant-78.github.io/portfolio/" -ForegroundColor White -BackgroundColor DarkGreen
} else {
    Write-Host "`nPush succeeded, but GitHub Pages may need manual enable in repo Settings > Pages." -ForegroundColor Yellow
}
Write-Host ""
