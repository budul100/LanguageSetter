param (
	[string] $changeLogPath,
	[string[]] $projectPaths
)

$scriptpath = split-path -parent $MyInvocation.MyCommand.Definition

. $scriptpath\Get_VersionCurrent.ps1

$currentVersion = Get-ProjectVersion -projectPath $projectPaths[0]

if ($currentVersion -cmatch '(?<major>\d+)\.(?<minor>\d+)\.(?<build>\d+)(\.(?<revision>\d+))?') 
{
	$updatedVersion = ($matches['major'] + "." + ([int]$matches['minor'] + 1) + ".0") 

	if (!($null -eq $matches['revision']))
	{
		$updatedVersion = $updatedVersion + ".0"
	}

} 
else 
{

	$updatedVersion = $currentVersion

}

. $scriptpath\Set_FileContent.ps1

foreach ($path in $projectPaths)
{
	Set-FileContent -path $path -replace '(?<=<AssemblyVersion>)[^<]+' -replaceWith "$updatedVersion"
}

Set-FileContent -path $changeLogPath -replace '# Change log' -replaceWith "# Change log`r`n`r`n## Version $updatedVersion`r`n`r`n..."

exit 0