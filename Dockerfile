FROM python:3.12-slim

# ffmpeg is required for audio compression before sending to Groq.
RUN apt-get update && \
    apt-get install -y --no-install-recommends ffmpeg && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install deps first for better Docker layer caching
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of the app
COPY . .

# Render injects PORT; default to 5000 for local docker runs
ENV PORT=5000
EXPOSE 5000

# Single worker is fine — this app is I/O bound on Groq + Claude.
# Long timeout because transcribing a 90-min recording can take ~30-60s.
CMD gunicorn -w 1 -b 0.0.0.0:${PORT} app:app --timeout 600 --access-logfile -
