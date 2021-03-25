CALL Unregister.bat
CALL Clean.bat

dotnet build "..\PrismTaskPanes\PrismTaskPanes.csproj"  --configuration Release
dotnet build "..\PrismTaskPanes.Regions\PrismTaskPanes.Regions.csproj"  --configuration Release
