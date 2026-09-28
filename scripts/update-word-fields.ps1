param([Parameter(Mandatory=$true)][string]$DocumentPath)
$ErrorActionPreference='Stop'
$resolved=(Resolve-Path -LiteralPath $DocumentPath).Path
$word=$null;$doc=$null
try {
  $word=New-Object -ComObject Word.Application
  $word.Visible=$false;$word.DisplayAlerts=0
  $doc=$word.Documents.Open($resolved,$false,$false)
  $doc.Repaginate()
  $doc.Fields.Update() | Out-Null
  foreach($toc in $doc.TablesOfContents){$toc.Update()}
  # Word 更新目录会在结尾保留一个空的域结束段，压缩它以防额外空白页。
  foreach($toc in $doc.TablesOfContents){
    foreach($p in $toc.Range.Paragraphs){
      if($p.Range.Text.Trim().Length -eq 0){
        $p.Range.Font.Size=1
        $p.Format.LineSpacingRule=4;$p.Format.LineSpacing=1
        $p.Format.SpaceBefore=0;$p.Format.SpaceAfter=0
        $p.Format.KeepWithNext=0;$p.Format.PageBreakBefore=0
      }
    }
  }
  $doc.Repaginate()
  foreach($toc in $doc.TablesOfContents){$toc.UpdatePageNumbers()}
  $doc.Save()
  Write-Output ('Word 已更新目录，共 '+$doc.ComputeStatistics(2)+' 页。')
} finally {
  if($doc){$doc.Close($false);[void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc)}
  if($word){$word.Quit();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)}
}
