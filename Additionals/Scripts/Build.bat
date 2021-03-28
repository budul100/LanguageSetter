CALL Unregister.bat
CALL Clean.bat

dotnet build "..\..\Main\Main.csproj"  --configuration Release
