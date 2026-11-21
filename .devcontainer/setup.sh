#!/usr/bin/env bash
# .devcontainer/setup.sh
set -e # Exit immediately if any command fails

sudo chsh -s /bin/bash $(whoami)

echo "=== Running postCreate Setup ==="

# Git identity. The container has no ~/.gitconfig of its own and the host's is not mounted, so
# without this every commit fails on "unable to auto-detect email address". Ahead of the installs
# below so a surprise here does not leave a half-built toolchain.
echo "Configuring Git identity ..."
git config --global user.name "pvillela"
git config --global user.email "pvillela@gmail.com"

# Local binary folder
export LOCAL_BIN="$HOME/.local/bin"
mkdir -p ${LOCAL_BIN}

# echo "Installing Rust ..."
# curl --proto '=https' --tlsv1.2 -sSfL https://sh.rustup.rs | bash

echo "Installing Claude Code..."

# Download and execute the official native installer (requires no sudo or Node at runtime)
curl -fsSL https://claude.ai/install.sh | bash

# Ensure ~/.local/bin is explicitly added to the PATH for non-interactive shells if needed
if [[ ":$PATH:" != *":$HOME/.local/bin:"* ]]; then
    echo 'export PATH="$HOME/.local/bin:$PATH"' >> ~/.bashrc
fi

echo "Claude Code installation complete!"

# echo "Installing Bun ..."
# curl -fsSL https://bun.com/install | bash
# export BUN_INSTALL="$HOME/.bun"
# export PATH="$BUN_INSTALL/bin:$PATH"

echo "Installing oh-my-pi ..."
# bun install -g @oh-my-pi/pi-coding-agent
curl -fsSL https://omp.sh/install | sh

# echo "Installing Gemini CLI ..."
# (source ${NVM_DIR}/nvm.sh && npm install -g @google/gemini-cli)

echo "Installing herdr ..."
curl -fsSL https://herdr.dev/install.sh | sh

echo "Installing Zellij ..."
ZELLIJ_VERSION=0.44.3
curl -LO https://github.com/zellij-org/zellij/releases/download/v${ZELLIJ_VERSION}/zellij-x86_64-unknown-linux-musl.tar.gz && \
    tar -xvzf zellij-x86_64-unknown-linux-musl.tar.gz && \
    chmod +x zellij
mv zellij ${LOCAL_BIN}/ && \
    rm zellij-x86_64-unknown-linux-musl.tar.gz

echo "Installing LibreOffice 26.2 ..."
# The workbook must work in LibreOffice Calc 26.2, and the migration and its checks drive Calc
# headless. The official packages are used rather than Ubuntu's, whose version may differ. They
# do not pull in their runtime libraries, hence the explicit list.
LO_VERSION=26.2.7
LO_TMP=$(mktemp -d)
curl -fL -o "${LO_TMP}/lo.tar.gz" \
    "https://download.documentfoundation.org/libreoffice/stable/${LO_VERSION}/deb/x86_64/LibreOffice_${LO_VERSION}_Linux_x86-64_deb.tar.gz"
tar -xzf "${LO_TMP}/lo.tar.gz" -C "${LO_TMP}"
sudo dpkg -i "${LO_TMP}"/LibreOffice_${LO_VERSION}*_deb/DEBS/*.deb
sudo apt-get update
sudo apt-get install -y libx11-xcb1 libxinerama1 libcairo2 libcups2t64 libdbus-1-3 libnspr4 libnss3
rm -rf "${LO_TMP}"

# Add local binary folder to PATH
echo "export PATH=\"\$LOCAL_BIN:\$PATH\"" >> "$HOME/.bashrc"

echo "=== Setup complete ==="