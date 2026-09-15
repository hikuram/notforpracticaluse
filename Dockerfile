FROM python:3.12-slim-trixie

ARG APP_VERSION=0.1.0

LABEL org.opencontainers.image.title="ppt2word-streamlit" \
      org.opencontainers.image.version="${APP_VERSION}" \
      org.opencontainers.image.description="Local Streamlit interface for converting PowerPoint/PDF materials into meeting-minutes DOCX" \
      org.opencontainers.image.authors="notforpracticaluse contributors" \
      org.opencontainers.image.licenses="MIT"

WORKDIR /app

ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    STREAMLIT_BROWSER_GATHER_USAGE_STATS=false \
    STREAMLIT_SERVER_HEADLESS=true \
    STREAMLIT_SERVER_ADDRESS=0.0.0.0 \
    STREAMLIT_SERVER_PORT=8501 \
    STREAMLIT_SERVER_FILE_WATCHER_TYPE=none \
    STREAMLIT_CLIENT_TOOLBAR_MODE=minimal \
    STREAMLIT_CLIENT_SHOW_ERROR_LINKS=false

RUN apt-get update \
    && apt-get install -y --no-install-recommends poppler-utils \
    && rm -rf /var/lib/apt/lists/*

RUN python -m pip install --no-cache-dir \
        "streamlit==1.56.0" \
        "lxml>=6.1.1" \
        "pdf2image>=1.17.0" \
        "Pillow>=12.3.0" \
        "python-docx>=1.2.0" \
        "python-pptx>=1.0.2" \
    && python -m pip check

COPY app.py ppt2word.py /app/
COPY templates /app/templates/

EXPOSE 8501

ENTRYPOINT ["streamlit", "run", "/app/app.py"]
