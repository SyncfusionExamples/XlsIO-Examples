#See https://aka.ms/customizecontainer to learn how to customize your debug container and how Visual Studio uses this Dockerfile to build your images for faster debugging.

FROM mcr.microsoft.com/dotnet/aspnet:3.1-buster-slim AS base
RUN apt-get update -y && apt-get install fontconfig -y
WORKDIR /app

FROM mcr.microsoft.com/dotnet/sdk:3.1-buster-slim AS build
WORKDIR /src
COPY ["Edit Excel.csproj", "."]
RUN dotnet restore "./Edit Excel.csproj"
COPY . .
WORKDIR "/src/."
RUN dotnet build "Edit Excel.csproj" -c Release -o /app/build

FROM build AS publish
RUN dotnet publish "Edit Excel.csproj" -c Release -o /app/publish

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "Edit Excel.dll"]