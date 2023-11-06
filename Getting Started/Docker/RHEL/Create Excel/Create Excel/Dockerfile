#See https://aka.ms/customizecontainer to learn how to customize your debug container and how Visual Studio uses this Dockerfile to build your images for faster debugging.

FROM registry.access.redhat.com/ubi8/dotnet-31-runtime AS base
USER root
RUN yum -y install fontconfig --disablerepo=epel
WORKDIR /

FROM registry.access.redhat.com/ubi8/dotnet-31 AS build
WORKDIR /src
COPY ["Create Excel.csproj", ""]
RUN dotnet restore "./Create Excel.csproj"
COPY . .
WORKDIR "/src/."
RUN dotnet build "Create Excel.csproj" -c Release -o /app/build

FROM build AS publish
RUN dotnet publish "Create Excel.csproj" -c Release -o /app/publish

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "Create Excel.dll"]