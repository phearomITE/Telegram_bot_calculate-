#!/bin/bash
echo "Starting bot with auto-restart on .py changes..."

watchmedo auto-restart \
  --directory=/app \
  --pattern="*.py" \
  --recursive \
  -- python -u bot.py   # replace with your main file
