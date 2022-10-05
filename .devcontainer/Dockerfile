FROM ubuntu:20.04

LABEL maintainer="Kasai Kou"
ARG apt_get_server=ftp.jaist.ac.jp/pub/Linux
ARG NODEJS_VERSION=14.x
ARG golang_version=1.18
ARG username=vscode
ARG useruid=1000
ARG usergid=${useruid}

ENV DEBIAN_FRONTEND=nointeractive
ENV LANG=C.UTF-8

ENV GOPATH=/home/${username}/go
ENV PATH=${PATH}:/usr/local/go/bin
ENV PATH=${PATH}:${GOPATH}/bin
ENV PATH=${PATH}:/home/${username}/.yarn/bin
ENV SHELL=/bin/bash
ENV CHROME_BIN=/usr/bin/chromium-browser
ENV PUPPETEER_SKIP_CHROMIUM_DOWNLOAD=true

WORKDIR /opt
RUN \
    sed -i s@archive.ubuntu.com@${apt_get_server}@g /etc/apt/sources.list && \
    apt-get update && \
    apt-get install -y --no-install-recommends \
    apt-transport-https \
    build-essential \
    curl \
    ca-certificates \
    ccache \
    lsb-release \
    pkg-config \
    gnupg \
    git \
    zip \
    unzip \
    python3-pip \
    python3-dev \
    patchelf \
    evince \
    texlive-latex-extra \
    texlive-fonts-extra \
    texlive-fonts-recommended \
    texlive-lang-cjk \
    xdvik-ja \
    dvipng \
    wget \
    nodejs \
    npm \
    sudo && \
    #
    # Install mermaid
    npm install --global yarn && \
    yarn global add --production puppeteer mermaid.cli && \
    wget https://dl.google.com/linux/direct/google-chrome-stable_current_amd64.deb && \
    apt-get install -y ./google-chrome-stable_current_amd64.deb && \
    #
    # Create non-root user: https://aka.ms/vscode-remote/containers/non-root-user
    groupadd --gid ${usergid} ${username} && \
    useradd --uid ${useruid} --gid ${usergid} -m ${username} && \
    echo ${username} ALL=\(root\) NOPASSWD:ALL > /etc/sudoers.d/${username} && \
    chmod 0440 /etc/sudoers.d/${username} && \
    chown -R ${username} /opt

USER ${username}
COPY requirements.lock .
RUN python3 -m pip install -r requirements.lock

RUN python3 -m pip install \
    mypy \
    yapf
