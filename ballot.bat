@echo off
setlocal enabledelayedexpansion

echo ============================================
echo   JCR Picture Ballot Assignment Engine
echo ============================================
echo.

:: Check if required files exist
if not exist "items.csv" (
    echo ERROR: items.csv not found!
    pause
    exit /b 1
)

if not exist "ballots.csv" (
    echo ERROR: ballots.csv not found!
    pause
    exit /b 1
)

:: Read seed from seed.txt file
if exist "seed.txt" (
    echo Reading seed from seed.txt...
    set /p USER_SEED=<seed.txt
    set "USER_SEED=!USER_SEED!"
    
    :: Clean up the seed (remove any carriage returns/line endings)
    set "USER_SEED=!USER_SEED: =!"
    set "USER_SEED=!USER_SEED:`r=!"
    set "USER_SEED=!USER_SEED:`n=!"
    
    echo Seed loaded from file: !USER_SEED!
    
    :: Optional: Ask if user wants to use a different seed
    choice /c YN /m "Use this seed? [Y/N]: "
    if errorlevel 2 (
        echo.
        set /p USER_SEED="Enter new seed: "
    )
) else (
    :: Create seed.txt with default seed if it doesn't exist
    echo 20260 > seed.txt
    echo Created seed.txt with default seed: 20260
    set /p USER_SEED=<seed.txt
)

:: Display seed being used
echo.
echo Using seed: !USER_SEED!
echo.

:: Ask for output prefix
set /p OUT_PREFIX="Enter output file prefix [output]: "
if "!OUT_PREFIX!"=="" set OUT_PREFIX=output

:: Ask for verbose mode
choice /c YN /m "Enable verbose output? [Y/N]: "
if errorlevel 2 (
    set VERBOSE_FLAG=
) else (
    set VERBOSE_FLAG=--verbose
)

:: Ask for max rounds override
set /p MAX_ROUNDS="Enter max rounds override [press Enter for default]: "
if not "!MAX_ROUNDS!"=="" (
    set MAX_ROUNDS_FLAG=--max-rounds !MAX_ROUNDS!
) else (
    set MAX_ROUNDS_FLAG=
)

echo.
echo ============================================
echo   Running with parameters:
echo   Seed: !USER_SEED!
echo   Items: items.csv
echo   Ballots: ballots.csv
echo   Output: !OUT_PREFIX!_!USER_SEED!_
echo   Verbose: !VERBOSE_FLAG!
echo   Max Rounds: !MAX_ROUNDS!
echo ============================================
echo.

:: Add seed to output prefix for better tracking
set FINAL_OUT_PREFIX=!OUT_PREFIX!_!USER_SEED!

:: Run the Python script
python3 ballot_assigner.py --items items.csv --ballots ballots.csv --seed !USER_SEED! --out !FINAL_OUT_PREFIX! !VERBOSE_FLAG! !MAX_ROUNDS_FLAG!

:: Check if Python script succeeded
if errorlevel 1 (
    echo.
    echo ERROR: Script failed with error code !errorlevel!
    pause
    exit /b !errorlevel!
)

echo.
echo ============================================
echo   Files created:
echo   - !FINAL_OUT_PREFIX!_allocations.csv
echo   - !FINAL_OUT_PREFIX!_audit_log.csv
echo ============================================
echo.

:: Option to open the output files
choice /c YN /m "Open output files? [Y/N]: "
if errorlevel 2 (
    echo Done. Seed used: !USER_SEED!
) else (
    echo Opening output files...
    start excel "!FINAL_OUT_PREFIX!_allocations.csv"
    timeout /t 1 >nul
    start excel "!FINAL_OUT_PREFIX!_audit_log.csv"
)

pause