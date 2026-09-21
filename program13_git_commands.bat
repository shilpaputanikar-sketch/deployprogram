@echo off
REM Git Version Control Commands
REM Run these commands in order

echo Step 1: Initialize Git repository
git init

echo.
echo Step 2: Check status
git status

echo.
echo Step 3: Add all files to staging
git add .

echo.
echo Step 4: Commit files
git commit -m "Initial commit: Backend API"

echo.
echo Step 5: Add remote origin
echo Create a repo on GitHub first, then run:
echo git remote add origin https://github.com/yourusername/your-repo.git

echo.
echo Step 6: Push to GitHub
echo git push -u origin main

echo.
echo Done. Use these commands anytime:
echo git status          - See changes
echo git add .           - Stage all changes
echo git commit -m "msg" - Commit changes
echo git push            - Upload to GitHub
