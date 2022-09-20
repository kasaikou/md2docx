FROM debian:bullseye

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
    wget \
    sudo && \
    #
    # Install docker
    curl -fsSL https://download.docker.com/linux/ubuntu/gpg | gpg --dearmor -o /usr/share/keyrings/docker-archive-keyring.gpg && \
    echo "deb [arch=$(dpkg --print-architecture) signed-by=/usr/share/keyrings/docker-archive-keyring.gpg] https://download.docker.com/linux/debian $(lsb_release -cs) stable" | tee /etc/apt/sources.list.d/docker.list > /dev/null && \
    apt-get update && \
    apt-get install -y docker-ce docker-ce-cli containerd.io && \
    #
    #
    wget https://go.dev/dl/go${golang_version}.linux-$(dpkg --print-architecture).tar.gz && \
    tar -C /usr/local -xzf go${golang_version}.linux-$(dpkg --print-architecture).tar.gz && \
    #
    # Create non-root user: https://aka.ms/vscode-remote/containers/non-root-user
    groupadd --gid ${usergid} ${username} && \
    useradd --uid ${useruid} --gid ${usergid} -m ${username} && \
    echo ${username} ALL=\(root\) NOPASSWD:ALL > /etc/sudoers.d/${username} && \
    chmod 0440 /etc/sudoers.d/${username} && \
    chown -R ${username} /opt

USER ${username}
RUN go install golang.org/x/tools/gopls@latest && \
    go install golang.org/x/lint/golint@latest && \
    go install github.com/go-delve/delve/cmd/dlv@master && \
    go install github.com/haya14busa/goplay/cmd/goplay@v1.0.0 && \
    go install github.com/fatih/gomodifytags@v1.16.0 && \
    go install github.com/josharian/impl@latest && \
    go install github.com/cweill/gotests/gotests@latest && \
    go install github.com/ramya-rao-a/go-outline@latest && \
    go install golang.org/x/tools/cmd/godoc@latest && \
    go install honnef.co/go/tools/cmd/staticcheck@latest
