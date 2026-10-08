@echo off
title MimicMail Multi-Instance Launcher
cls
echo ===================================================================
echo             MimicMail - Parallel Multi-Instance Launcher
echo ===================================================================
echo.
echo Launching 2 parallel instances with isolated sessions...
echo.

python launch_parallel.py -n 2

echo.
echo Instances launched! You can close this window at any time.
pause
