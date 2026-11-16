param([Parameter(Mandatory=$true)][string]$InputPath,
      [Parameter(Mandatory=$true)][string]$RequestPath,
      [Parameter(Mandatory=$true)][string]$OutputDir)
$ErrorActionPreference='Stop'
$mutex=New-Object System.Threading.Mutex($false,'Local\ppt-again-chart-export')
$held=$false; $app=$null; $deck=$null; $originalSecurity=$null
$existing=@(Get-Process POWERPNT -ErrorAction SilentlyContinue)
function Find-Chart($shapes,[int]$id){
  foreach($shape in $shapes){
    if($shape.Id -eq $id){return $shape}
    if($shape.Type -eq 6){$found=Find-Chart $shape.GroupItems $id;if($found){return $found}}
  }
  return $null
}
try {
  $held=$mutex.WaitOne(120000);if(!$held){throw 'Chart exporter is busy'}
  $requests=ConvertFrom-Json -InputObject (Get-Content -LiteralPath $RequestPath -Raw -Encoding UTF8)
  if($requests -isnot [array]){$requests=@($requests)}
  try {$app=New-Object -ComObject PowerPoint.Application} catch {throw 'PPT_AGAIN_POWERPOINT_UNAVAILABLE'}
  try {$originalSecurity=$app.AutomationSecurity;$app.AutomationSecurity=3;if($app.AutomationSecurity -ne 3){throw 'security'}} catch {throw 'PPT_AGAIN_POWERPOINT_SECURITY_FAILED'}
  try {$deck=$app.Presentations.Open($InputPath,-1,0,0)} catch {throw 'PPT_AGAIN_POWERPOINT_OPEN_FAILED'}
  if($app.AutomationSecurity -eq 3){$app.AutomationSecurity=$originalSecurity}
  $originalSecurity=$null
  $results=@()
  for($i=0;$i -lt $requests.Count;$i++){
    $item=$requests[$i];$slideNumber=[int]$item.slideNumber;$shapeId=[int]$item.shapeId
    $result=@{slideNumber=$slideNumber;shapeId=$shapeId}
    try {
      if($slideNumber -lt 1 -or $slideNumber -gt $deck.Slides.Count -or $shapeId -lt 1){throw 'PPT_AGAIN_CHART_NOT_FOUND'}
      $shape=Find-Chart $deck.Slides.Item($slideNumber).Shapes $shapeId
      if(!$shape -or $shape.HasChart -ne -1){throw 'PPT_AGAIN_CHART_NOT_FOUND'}
      # Shape.Export scales against the slide dimensions, not the shape bounds.
      $slideWidth=$deck.PageSetup.SlideWidth;$slideHeight=$deck.PageSetup.SlideHeight
      $scale=[math]::Min(2,4096/[math]::Max($slideWidth,$slideHeight))
      $width=[int][math]::Max(1,[math]::Round($slideWidth*$scale))
      $height=[int][math]::Max(1,[math]::Round($slideHeight*$scale))
      $file=Join-Path $OutputDir ("chart-{0}.png" -f ($i+1))
      $shape.Export($file,2,$width,$height,1)
      if(!(Test-Path -LiteralPath $file -PathType Leaf)){throw 'Chart export did not create a file'}
    } catch {$result.errorCode=if($_.Exception.Message -eq 'PPT_AGAIN_CHART_NOT_FOUND'){'POWERPOINT_CHART_NOT_FOUND'}else{'POWERPOINT_EXPORT_FAILED'}}
    $results+= $result
  }
  $json=ConvertTo-Json -InputObject @($results) -Compress
  [IO.File]::WriteAllText((Join-Path $OutputDir 'result.json'),$json,(New-Object Text.UTF8Encoding($false)))
} finally {
  try {if($deck){try{$deck.Close()}finally{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($deck)}}}
  finally {try {if($app){try {if($null -ne $originalSecurity -and $app.AutomationSecurity -eq 3){$app.AutomationSecurity=$originalSecurity};if($existing.Count -eq 0 -and $app.Presentations.Count -eq 0){$app.Quit()}} finally {[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}}}
    finally {if($held){$mutex.ReleaseMutex()};$mutex.Dispose()}}
}
