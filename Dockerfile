# ==============================================================================
# PowerPoint to Word "Total Victory V4" Container (ONLYOFFICE Edition)
# ==============================================================================
# Use the latest Debian Bookworm-based Python slim image
FROM python:3.12-slim-bookworm

# Disable interactive prompts during package installation
ENV DEBIAN_FRONTEND=noninteractive

# 1. Install required packages and OSS Japanese fonts
# (Using --no-install-recommends to exclude unnecessary dependencies and keep the image small)
RUN apt-get update && apt-get install -y --no-install-recommends \
    wget \
    curl \
    gnupg \
    poppler-utils \
    fonts-ipafont \
    fonts-ipaexfont \
    fontconfig \
    && rm -rf /var/lib/apt/lists/*

# 2. Add ONLYOFFICE repository and install Document Builder
# (Fetching the GPG key directly from the official source for better reliability)
RUN curl -fsSL https://download.onlyoffice.com/repo/onlyoffice.asc | gpg --dearmor > /etc/apt/trusted.gpg.d/onlyoffice.gpg \
    && echo "deb https://download.onlyoffice.com/repo/debian squeeze main" | tee /etc/apt/sources.list.d/onlyoffice.list \
    && apt-get update \
    && apt-get install -y --no-install-recommends onlyoffice-documentbuilder \
    && rm -rf /var/lib/apt/lists/*

# 3. Install Python dependencies
# (Using the built-in pip provided by the official Python image)
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir \
    python-pptx \
    python-docx \
    Pillow \
    pdf2image

# Set the working directory
WORKDIR /workspace

# Set bash as the default command
CMD ["/bin/bash"]
