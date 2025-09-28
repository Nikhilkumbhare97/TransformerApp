# Use the official .NET 8 runtime as base image
FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS base
WORKDIR /app
EXPOSE 80
EXPOSE 443

# Use the SDK image to build the application
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src

# Copy project file and restore dependencies
COPY ["InventorApp.API.csproj", "."]
RUN dotnet restore "InventorApp.API.csproj"

# Copy the rest of the source code
COPY . .

# Build the application
RUN dotnet build "InventorApp.API.csproj" -c Release -o /app/build

# Publish the application
FROM build AS publish
RUN dotnet publish "InventorApp.API.csproj" -c Release -o /app/publish /p:UseAppHost=false

# Create the final runtime image
FROM base AS final
WORKDIR /app

# Copy the published application
COPY --from=publish /app/publish .

# Copy the Jacob library for Inventor integration
COPY --from=build /src/libs/jacob.jar ./libs/

# Set environment variables
ENV ASPNETCORE_URLS=http://+:80
ENV ASPNETCORE_ENVIRONMENT=Production

# Create a non-root user
RUN adduser --disabled-password --gecos '' appuser && chown -R appuser /app
USER appuser

ENTRYPOINT ["dotnet", "InventorApp.API.dll"]
