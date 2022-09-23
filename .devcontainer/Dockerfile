FROM ubuntu:20.04

LABEL maintainer="Kasai Kou"

ARG apt_get_server=ftp.jaist.ac.jp/pub/Linux
ARG golang_version=1.18
ARG username=vscode
ARG useruid=1000
ARG usergid=${useruid}

ENV DEBIAN_FRONTEND=nointeractive
ENV LANG=C.UTF-8

ENV GOPATH=/home/${username}/go
ENV PATH=${PATH}:/usr/local/go/bin
ENV PATH=${PATH}:${GOPATH}/bin
ENV SHELL=/bin/bash

WORKDIR /opt
RUN \
    sed -i s@archive.ubuntu.com@${apt_get_server}@g /etc/apt/sources.list && \
    apt-get update && \
    apt-get install -y --no-install-recommends \
    apt-transport-https \
    build-essential \
    curl \
    ca-certificates \
    lsb-release \
    pkg-config \
    gnupg \
    git \
    zip \
    unzip \
    python3-pip \
    evince \
    texlive-fonts-extra \
    texlive-fonts-recommended \
    texlive-lang-cjk \
    xdvik-ja \
    dvipng \
    wget \
    sudo && \
    #
    # Install docker
    curl -fsSL https://download.docker.com/linux/ubuntu/gpg | gpg --dearmor -o /usr/share/keyrings/docker-archive-keyring.gpg && \
    echo "deb [arch=$(dpkg --print-architecture) signed-by=/usr/share/keyrings/docker-archive-keyring.gpg] https://download.docker.com/linux/ubuntu $(lsb_release -cs) stable" | tee /etc/apt/sources.list.d/docker.list > /dev/null && \
    apt-get update && \
    apt-get install -y docker-ce docker-ce-cli containerd.io && \
    #
    # Create non-root user: https://aka.ms/vscode-remote/containers/non-root-user
    groupadd --gid ${usergid} ${username} && \
    useradd --uid ${useruid} --gid ${usergid} -m ${username} && \
    echo ${username} ALL=\(root\) NOPASSWD:ALL > /etc/sudoers.d/${username} && \
    chmod 0440 /etc/sudoers.d/${username} && \
    chown -R ${username} /opt

USER ${username}
RUN python3 -m pip install \
    mypy \
    yapf

