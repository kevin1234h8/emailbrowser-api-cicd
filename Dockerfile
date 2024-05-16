FROM mcr.microsoft.com/dotnet/sdk:7.0 AS build-env
WORKDIR /app

# Copy the csproj and restore dependencies
COPY *.csproj ./
RUN dotnet restore

# Copy the rest of the application code
COPY . ./

# Build the application and publish it to the /out directory
RUN dotnet publish -c Release -o /out

# Use the official .NET runtime image to run the application
FROM mcr.microsoft.com/dotnet/aspnet:7.0
WORKDIR /app

# Copy the published application from the build environment
COPY --from=build-env /out .

# Expose the port on which the application will run
EXPOSE 80

# Define the entry point for the container to run the application
ENTRYPOINT ["dotnet", "MyWebApi.dll"]