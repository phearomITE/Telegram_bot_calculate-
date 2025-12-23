FROM python:3.11-slim

WORKDIR /app

# Install Python deps
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy project files
COPY . .

# For development: run file-watcher which restarts bot on .py changes
CMD ["python", "run_bot.py"]
