FROM python:3.11-slim

WORKDIR /app

# Install Chinese fonts (宋体 etc.) for docx generation
RUN apt-get update && \
    apt-get install -y --no-install-recommends fonts-wqy-zenhei fonts-noto-cjk && \
    rm -rf /var/lib/apt/lists/*

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code and assets
COPY report_tool.py section_builders.py index.html ./
COPY 模版_cover.docx 模版.docx ./

# Output directory (replaces ~/Desktop/gradio_report_jobs inside container)
ENV JOB_ROOT=/app/output
RUN mkdir -p /app/output

# Expose Flask port
EXPOSE 7860

# Run with production WSGI server
CMD ["python", "report_tool.py"]
