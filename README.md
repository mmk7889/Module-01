Get-AppxPackage *windowscalculator* | Foreach {Add-AppxPackage -DisableDevelopmentMode -Register "$($_.InstallLocation)\AppXManifest.xml"}
9WZDNCRFHVN5
