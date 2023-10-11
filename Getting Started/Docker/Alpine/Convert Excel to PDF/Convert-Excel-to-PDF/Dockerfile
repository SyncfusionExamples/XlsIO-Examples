#See https://aka.ms/customizecontainer to learn how to customize your debug container and how Visual Studio uses this Dockerfile to build your images for faster debugging.

FROM mcr.microsoft.com/dotnet/aspnet:3.1-alpine3.12 AS base
RUN apk update && apk upgrade && apk add fontconfig
RUN apk add --update ttf-dejavu fontconfig
WORKDIR /app

FROM mcr.microsoft.com/dotnet/sdk:3.1-alpine3.12 AS build
WORKDIR /src
COPY ["Convert-Excel-to-PDF.csproj", "."]
RUN dotnet restore "./Convert-Excel-to-PDF.csproj"
COPY . .
WORKDIR "/src/."
RUN dotnet build "Convert-Excel-to-PDF.csproj" -c Release -o /app/build

FROM build AS publish
RUN dotnet publish "Convert-Excel-to-PDF.csproj" -c Release -o /app/publish

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "Convert-Excel-to-PDF.dll"]