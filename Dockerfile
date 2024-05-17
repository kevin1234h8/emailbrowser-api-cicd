FROM mcr.microsoft.com/dotnet/sdk:7.0 AS build-env
WORKDIR /source
COPY . .
RUN dotnet restore
RUN dotnet publish -c Release -o/app --no-restore
FROM mcr.microsoft.com/dotnet/aspnet:7.0
WORKDIR /app
COPY --from=build-env /out .
EXPOSE 5000

# Define the entry point for the container to run the application
ENTRYPOINT ["dotnet", "AGOServer.dll"]