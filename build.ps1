Add-Type -AssemblyName System.IO.Compression.FileSystem
if (-not (Test-Path .\nuget.exe)) {
    Invoke-WebRequest -Uri "https://dist.nuget.org/win-x86-commandline/latest/nuget.exe" -OutFile nuget.exe
}

$ikvm_version = '8.1.5717.0'

if (-not(Test-Path ikvm\)) {
    Invoke-WebRequest -Uri "http://www.frijters.net/ikvmbin-$ikvm_version.zip" -OutFile ikvm.zip -UserAgent "NativeHost"
    [System.IO.Compression.ZipFile]::ExtractToDirectory("$PWD\ikvm.zip", "$PWD")
    Remove-Item ikvm.zip
    Get-Item ikvm* | Select-Object -Index 0 | Rename-Item -Path { $_.FullName } -NewName ikvm
}

if (-not(Test-Path poi\)) {
    $poi_last_supported_url = "https://archive.apache.org/dist/poi/release/bin/poi-bin-4.1.0-20190412.zip"

    Invoke-WebRequest -Uri $poi_last_supported_url -OutFile poi.zip

    [System.IO.Compression.ZipFile]::ExtractToDirectory("$PWD\poi.zip", "$PWD")

    Remove-Item poi.zip

    Get-Item poi* | Select-Object -Index 0 | Rename-Item -Path { $_.FullName } -NewName poi
}

$poi_common = Get-ChildItem -Path poi | Where-Object { $_.Name -match '^poi-[.\d]+\.jar$' }
$poi_scratchpad = Get-ChildItem -Path poi | Where-Object { $_.Name -match '^poi-scratchpad-[.\d]+\.jar$' }
$poi_ooxml = Get-ChildItem -Path poi | Where-Object { $_.Name -match '^poi-ooxml-[.\d]+\.jar$' }

$version = [Version]$poi_common.BaseName.Substring("poi-".Length)
$version = New-Object -TypeName Version -ArgumentList ($version.Major, [Math]::Max(0, $version.Minor), [Math]::Max(0, $version.Build), [Math]::Max(0, $version.Revision))

$latestPublished = (.\nuget.exe list id:POI-IKVM.Fork) | Where-Object { $_ -imatch '^POI-IKVM.Fork\s' } | ForEach-Object { [Version]($_ -replace '^POI-IKVM.Fork\s(.*)$', '$1') }
$latestPublished = New-Object -TypeName Version -ArgumentList ($latestPublished.Major, [Math]::Max(0, $latestPublished.Minor), [Math]::Max(0, $latestPublished.Build), [Math]::Max(0, $latestPublished.Revision))

if ($null -ne $latestPublished -and $latestPublished -ge $version) {
    $version = New-Object -TypeName Version -ArgumentList ($latestPublished.Major, $latestPublished.Minor, $latestPublished.Build, ($latestPublished.Revision + 1))
}

$vr = '[0-9.\-]+'

$libs = "activation$($vr)jar|commons-[a-z]+$($vr)jar|jaxb-[a-z]+$($vr)jar|junit$($vr)jar|log4j$($vr)jar|SparseBitSet$($vr)jar"

$new_libs = Get-ChildItem -Path .\poi\lib | Select-Object { $_.Name -imatch $libs } | Where-Object { $_ -eq $false }

if ($null -ne $new_libs) {
    throw "This version of poi has new dependencies. The code below needs to be modified to determine whether to include them in the package"
}

$poiCommonJars = @($($poi_common; (Get-ChildItem -Path .\poi\lib | Where-Object { $_ -inotmatch "activation$($vr)jar|junit$($vr)jar|jaxb-[a-z]+$($vr)jar" })) | Select-Object -ExpandProperty FullName)
$poiScratchpadJars = @($($poi_common; $poi_scratchpad; (Get-ChildItem -Path .\poi\lib | Where-Object { $_ -inotmatch "activation$($vr)jar|junit$($vr)jar|jaxb-[a-z]+$($vr)jar" })) | Select-Object -ExpandProperty FullName)
$poiOoxmlJars = @($($poi_common; $poi_ooxml; (Get-ChildItem -Path .\poi\lib | Where-Object { $_ -inotmatch "activation$($vr)jar|junit$($vr)jar|jaxb-[a-z]+$($vr)jar" })) | Select-Object -ExpandProperty FullName)

New-Item -Path .\packages\POI-IKVM.Fork\lib\net40 -ItemType Directory -Force

&.\ikvm\bin\ikvmc.exe -keyfile:key.snk -out:packages\POI-IKVM.Fork\lib\net40\POI-IKVM.Fork.dll -target:library ("-version:" + $version.ToString()) ("-fileversion:" + $version.ToString()) $poiCommonJars

$versionElement = Select-Xml -Path .\packages\POI-IKVM.Fork\POI-IKVM.Fork.nuspec -XPath 'package/metadata/version'
$versionElement.Node.'#text' = $version.ToString()
$versionElement.Node.OwnerDocument.Save("$PWD\packages\POI-IKVM.Fork\POI-IKVM.Fork.nuspec")

$versionElement = Select-Xml -Path .\packages\POI-IKVM.Fork\POI-IKVM.Fork.nuspec -XPath 'package/metadata/dependencies/group/dependency[@id=''IKVM'']'
$versionElement.Node.SetAttribute('version', $ikvm_version)
$versionElement.Node.OwnerDocument.Save("$PWD\packages\POI-IKVM.Fork\POI-IKVM.Fork.nuspec")

New-Item -Path .\packages\POI-IKVM.Scratchpad.Fork\lib\net40 -ItemType Directory -Force

&.\ikvm\bin\ikvmc.exe -keyfile:key.snk -out:packages\POI-IKVM.Scratchpad.Fork\lib\net40\POI-IKVM.Scratchpad.Fork.dll -target:library ("-version:" + $version.ToString()) ("-fileversion:" + $version.ToString()) $poiScratchpadJars

$versionElement = Select-Xml -Path .\packages\POI-IKVM.Scratchpad.Fork\POI-IKVM.Scratchpad.Fork.nuspec -XPath 'package/metadata/version'
$versionElement.Node.'#text' = $version.ToString()
$versionElement.Node.OwnerDocument.Save("$PWD\packages\POI-IKVM.Scratchpad.Fork\POI-IKVM.Scratchpad.Fork.nuspec")

$versionElement = Select-Xml -Path .\packages\POI-IKVM.Scratchpad.Fork\POI-IKVM.Scratchpad.Fork.nuspec -XPath 'package/metadata/dependencies/group/dependency[@id=''IKVM'']'
$versionElement.Node.SetAttribute('version', $ikvm_version)
$versionElement.Node.OwnerDocument.Save("$PWD\packages\POI-IKVM.Scratchpad.Fork\POI-IKVM.Scratchpad.Fork.nuspec")

New-Item -Path .\packages\POI-IKVM.OOXML.Fork\lib\net40 -ItemType Directory -Force

&.\ikvm\bin\ikvmc.exe -keyfile:key.snk -out:packages\POI-IKVM.OOXML.Fork\lib\net40\POI-IKVM.OOXML.Fork.dll -target:library ("-version:" + $version.ToString()) ("-fileversion:" + $version.ToString()) $poiOoxmlJars

$versionElement = Select-Xml -Path .\packages\POI-IKVM.OOXML.Fork\POI-IKVM.OOXML.Fork.nuspec -XPath 'package/metadata/version'
$versionElement.Node.'#text' = $version.ToString()
$versionElement.Node.OwnerDocument.Save("$PWD\packages\POI-IKVM.OOXML.Fork\POI-IKVM.OOXML.Fork.nuspec")

$versionElement = Select-Xml -Path .\packages\POI-IKVM.OOXML.Fork\POI-IKVM.OOXML.Fork.nuspec -XPath 'package/metadata/dependencies/group/dependency[@id=''IKVM'']'
$versionElement.Node.SetAttribute('version', $ikvm_version)
$versionElement.Node.OwnerDocument.Save("$PWD\packages\POI-IKVM.OOXML.Fork\POI-IKVM.OOXML.Fork.nuspec")
