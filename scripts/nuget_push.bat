set nupkgname=%1%
set key=%2%

dotnet nuget push %nupkgname% -k %key% -s 
