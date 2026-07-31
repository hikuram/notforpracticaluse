# ==============================================================================
# PowerPoint to Word "Ninja" Container (ONLYOFFICE Edition)
# ==============================================================================
FROM ubuntu:22.04

# 対話モードを無効化
ENV DEBIAN_FRONTEND=noninteractive

# 1. 必須パッケージとフォントのインストール
# （IPAフォントなどのOSS日本語フォントをデフォルトで入れる）
RUN apt-get update && apt-get install -y \
    wget \
    gnupg \
    curl \
    python3 \
    python3-pip \
    poppler-utils \
    fonts-ipafont \
    fonts-ipaexfont \
    fontconfig \
    && rm -rf /var/lib/apt/lists/*

# 2. ONLYOFFICE Document Builder のリポジトリ追加とインストール
RUN mkdir -p -m 700 ~/.gnupg \
    && gpg --no-default-keyring --keyring gnupg-ring:/etc/apt/trusted.gpg.d/onlyoffice.gpg --keyserver hkp://keyserver.ubuntu.com:80 --recv-keys CB2DE8E5 \
    && chmod 644 /etc/apt/trusted.gpg.d/onlyoffice.gpg \
    && echo "deb https://download.onlyoffice.com/repo/debian squeeze main" | tee /etc/apt/sources.list.d/onlyoffice.list \
    && apt-get update \
    && apt-get install -y onlyoffice-documentbuilder \
    && rm -rf /var/lib/apt/lists/*

# 3. Pythonライブラリのインストール
RUN pip3 install --no-cache-dir \
    python-pptx \
    python-docx \
    Pillow \
    pdf2image

# 作業ディレクトリの設定
WORKDIR /workspace

# デフォルトのコマンドをbashに設定
CMD ["/bin/bash"]
