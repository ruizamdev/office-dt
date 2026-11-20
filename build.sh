#!/usr/bin/env sh
# Compila InstaladorOffice.exe (.NET Framework 4.8) desde Linux o macOS con el SDK de .NET.
# El exe queda en dist/.
set -e
cd "$(dirname "$0")"
dotnet build src/InstaladorOffice/InstaladorOffice.csproj -c Release -nologo -o dist
rm -f dist/*.config
echo "Listo: $(pwd)/dist/InstaladorOffice.exe"
