FROM mcr.microsoft.com/dotnet/runtime:5.0 AS base
WORKDIR /app

FROM mcr.microsoft.com/dotnet/sdk:5.0 AS build
WORKDIR /src
COPY ["CreateSpreadSheet/CreateSpreadSheet.csproj", "CreateSpreadSheet/"]
RUN dotnet restore "CreateSpreadSheet/CreateSpreadSheet.csproj"
COPY . .
WORKDIR "/src/CreateSpreadSheet"
RUN dotnet build "CreateSpreadSheet.csproj" -c Release -o /app/build

FROM build AS publish
RUN dotnet publish "CreateSpreadSheet.csproj" -c Release -o /app/publish

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "CreateSpreadSheet.dll"]
