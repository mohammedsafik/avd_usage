FROM mcr.microsoft.com/powershell:7.5-ubuntu-24.04

# Install prerequisites
RUN apt-get update && \
    apt-get install -y curl ca-certificates apt-transport-https gnupg lsb-release && \
    \
    # Install Node.js 22
    curl -fsSL https://deb.nodesource.com/setup_22.x | bash - && \
    apt-get install -y nodejs && \
    \
    # Install Azure CLI
    curl -sL https://aka.ms/InstallAzureCLIDeb | bash && \
    \
    # Optional: install exceljs globally
    npm install -g exceljs && \
    \
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app

CMD ["pwsh"]