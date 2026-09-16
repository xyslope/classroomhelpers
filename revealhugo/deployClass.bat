@echo off
chcp 65001 > nul
REM === Usage check ===
if "%1"=="--help" (
    echo Usage: deployClass.bat [--no-push]
    echo.
    echo Options:
    echo   --no-push    Build and commit locally without pushing to GitHub
    echo.
    exit /b 0
)
REM === Configuration ===
SET REPO_PATH=%CD%
SET WORK_BRANCH=no_notes
SET PUSH_TO_GITHUB=yes
if "%1"=="--no-push" SET PUSH_TO_GITHUB=no
echo ========================================
echo   Deploy Script for Class Materials
echo ========================================
echo.
REM === Step 1: Commit current state ===
echo [1/7] Committing current state on main branch...
cd /d %REPO_PATH%
git add -A
git commit -m "Save current state before deployment" >nul 2>&1
if errorlevel 1 (
    echo [INFO] No changes to commit on main branch
) else (
    echo [OK] Committed current state
)
REM === Step 2: Create no_notes branch ===
echo.
echo [2/7] Creating %WORK_BRANCH% branch...
git branch -D %WORK_BRANCH% 2>nul
git checkout -b %WORK_BRANCH%
if errorlevel 1 (
    echo [ERROR] Failed to create branch
    exit /b 1
)
echo [OK] Switched to %WORK_BRANCH% branch
REM === Step 3: Remove notes and Q&A ===
echo.
echo [3/7] Removing speaker notes and QA from Markdown...
python C:\Users\yusakata\work\github.com\xyslope\classroomhelpers\revealhugo\postEditContents.py
if errorlevel 1 (
    echo [ERROR] Failed to remove notes and QA
    git checkout main
    exit /b 1
)
REM === Step 4: Commit Markdown changes ===
echo.
echo [4/7] Committing Markdown changes...
git add -A
git commit -m "Remove speaker notes and Q&A slides"
if errorlevel 1 (
    echo [INFO] No Markdown changes to commit
) else (
    echo [OK] Committed Markdown changes
)
REM === Step 5: Build with Hugo ===
echo.
echo [5/7] Building site with Hugo...
hugo --cleanDestinationDir
if errorlevel 1 (
    echo [ERROR] Hugo build failed
    git checkout main
    exit /b 1
)
echo [OK] Hugo build completed
REM === Step 6: Commit build output ===
echo.
echo [6/7] Committing build output...
git add -A
git commit -m "Add Hugo build output"
if errorlevel 1 (
    echo [INFO] No build output changes
) else (
    echo [OK] Committed build output
)
REM === Step 7: Push to GitHub (optional) ===
echo.
if "%PUSH_TO_GITHUB%"=="yes" (
    echo [7/7] Pushing to GitHub from public folder...
    cd /d %REPO_PATH%\public
    git add -A
    git commit -m "Deploy from no_notes branch" >nul 2>&1
    git push origin main --force
    if errorlevel 1 (
        echo [ERROR] Failed to push to GitHub
        cd /d %REPO_PATH%
        git checkout main
        exit /b 1
    )
    echo [OK] Pushed to GitHub
    cd /d %REPO_PATH%
) else (
    echo [7/7] Skipping GitHub push
)
REM === Return to main branch ===
echo.
echo [CLEANUP] Returning to main branch...
git checkout main
if errorlevel 1 (
    echo [ERROR] Failed to return to main branch
    exit /b 1
)
echo.
echo ========================================
echo   Deployment completed successfully
echo ========================================
echo.
echo Branch "%WORK_BRANCH%" is ready
if "%PUSH_TO_GITHUB%"=="yes" (
    echo Published to GitHub
) else (
    echo Local only - not pushed to GitHub
    echo To push later: cd public and git push origin main --force
)
echo.
exit /b 0
