import os

# Try to read BOT_TOKEN from environment.
# If not set, fall back to the hard-coded token.
BOT_TOKEN = os.getenv("BOT_TOKEN") or "8420018950:AAFZugWzWORp4jJLi3aJZE-0Kw4jV7r9Vbg"

# Default exchange rate used in your calculations
EXCHANGE_RATE_DEFAULT = 4000
