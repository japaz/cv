FROM ubuntu:latest

# Avoid prompts during package installation
ENV DEBIAN_FRONTEND=noninteractive

# Install dependencies
RUN apt-get update && \
    echo "ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true" | debconf-set-selections && \
    apt-get install -y \
    pandoc \
    texlive-xetex \
    ttf-mscorefonts-installer \
    fontconfig \
    git \
    bash \
    --no-install-recommends && \
    fc-cache -fv && \
    rm -rf /var/lib/apt/lists/*

# Create working directory
WORKDIR /app

# Copy project files
COPY . .

# Ensure script is executable
RUN chmod +x process.sh

# Command to run the script
CMD ["./process.sh"]
