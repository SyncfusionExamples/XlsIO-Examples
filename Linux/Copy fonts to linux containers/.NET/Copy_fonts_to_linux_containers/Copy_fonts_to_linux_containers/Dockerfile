# Base runtime image (root user by default)
FROM mcr.microsoft.com/dotnet/runtime:8.0 AS base

# Install native deps needed by SkiaSharp/XlsIO renderer
RUN apt-get update && apt-get install -y --no-install-recommends \
    fontconfig \
    libfreetype6 \
    libpng16-16 \
    libexpat1 \
    libuuid1 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy fonts from your project folder into the container
COPY ["Copy_fonts_to_linux_containers/Fonts/", "/usr/local/share/fonts/"]

# Refresh the font cache so the new fonts are recognized
RUN fc-cache -f -v

# ------------------ build stage ------------------
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
ARG BUILD_CONFIGURATION=Release
WORKDIR /src

COPY ["Copy_fonts_to_linux_containers/Copy_fonts_to_linux_containers.csproj", "Copy_fonts_to_linux_containers/"]
RUN dotnet restore "./Copy_fonts_to_linux_containers/Copy_fonts_to_linux_containers.csproj"

COPY . .
WORKDIR "/src/Copy_fonts_to_linux_containers"
RUN dotnet build "./Copy_fonts_to_linux_containers.csproj" -c $BUILD_CONFIGURATION -o /app/build

# ------------------ publish stage ------------------
FROM build AS publish
ARG BUILD_CONFIGURATION=Release
RUN dotnet publish "./Copy_fonts_to_linux_containers.csproj" -c $BUILD_CONFIGURATION -o /app/publish /p:UseAppHost=false

# ------------------ final stage ------------------
FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "Copy_fonts_to_linux_containers.dll"]