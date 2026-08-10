# PowerPoint-to-Word processing container.
# PDF rendering is performed on the Windows host with Microsoft PowerPoint.
FROM python:3.12-slim-bookworm

ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

RUN apt-get update \
    && apt-get install -y --no-install-recommends poppler-utils \
    && rm -rf /var/lib/apt/lists/*

# These are minimum tested versions. Newer compatible releases are allowed.
RUN python -m pip install --no-cache-dir "pip>=26.2" \
    && python -m pip install --no-cache-dir \
        "lxml>=6.1.1" \
        "pdf2image>=1.17.0" \
        "Pillow>=12.3.0" \
        "python-docx>=1.2.0" \
        "python-pptx>=1.0.2"

WORKDIR /workspace

COPY ppt2word.py /opt/ppt2word/ppt2word.py
COPY header_template.docx /opt/ppt2word/header_template.docx

ENTRYPOINT ["python3", "/opt/ppt2word/ppt2word.py"]
