# Reconfigure la source locale apres deplacement du projet. Fermer Power BI avant execution.
$ErrorActionPreference='Stop'
$path=Join-Path $PSScriptRoot 'PowerBI\DataCo.SemanticModel\definition\expressions.tmdl'
if(-not (Test-Path -LiteralPath $path -PathType Leaf)){throw 'Fichier expressions.tmdl introuvable.'}
$folder=(Join-Path $PSScriptRoot 'exports').Replace('\','/')
if(-not (Test-Path -LiteralPath (Join-Path $PSScriptRoot 'exports') -PathType Container)){throw 'Dossier exports introuvable.'}
$text=[IO.File]::ReadAllText($path)
$replacement='expression FolderExports = "'+$folder.Replace('"','""')+'" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]'
$pattern='(?m)^expression FolderExports\s*=.*$'
if(-not [regex]::IsMatch($text,$pattern)){throw 'Parametre FolderExports absent.'}
$text=[regex]::Replace($text,$pattern,[System.Text.RegularExpressions.MatchEvaluator]{param($m) $replacement})
[IO.File]::WriteAllText($path,$text,[Text.UTF8Encoding]::new($false))
Write-Output ('Source Power BI configuree : '+$folder)
