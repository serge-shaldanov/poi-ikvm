if (Test-Path .\nupkgs) { Remove-Item .\nupkgs -Recurse -Force }
.\nuget.exe pack .\packages\POI-IKVM.Fork\POI-IKVM.Fork.nuspec -OutputDirectory .\nupkgs
.\nuget.exe pack .\packages\POI-IKVM.Scratchpad.Fork\POI-IKVM.Scratchpad.Fork.nuspec -OutputDirectory .\nupkgs
.\nuget.exe pack .\packages\POI-IKVM.OOXML.Fork\POI-IKVM.OOXML.Fork.nuspec -OutputDirectory .\nupkgs
.\nuget.exe push .\nupkgs\*.nupkg -Source https://www.nuget.org/api/v2/package
