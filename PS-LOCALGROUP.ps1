$csvpath = "c:\data\"
$csvoutpath = "c:\data\out.csv"

foreach ($group in Get-LocalGroup) {
    $localGroup =  [PSCustomObject]@{
                                        GroupName=($group.Name  | Out-String).Trim()
                                        Description=($group.Description  | Out-String).Trim()
                                     }
    $localGroup | Export-CSV $csvpath"Local-group.csv" -Append -NoTypeInformation -Force
}
   $localGroup =  [PSCustomObject]@{}
    $localGroup | Export-CSV $csvpath"Local-group.csv" -Append -NoTypeInformation -Force
     $localGroup | Export-CSV $csvpath"Local-group.csv" -Append -NoTypeInformation -Force
      $localGroup | Export-CSV $csvpath"Local-group.csv" -Append -NoTypeInformation -Force

foreach ($group in Get-LocalGroup) {

    $groupMember = Get-LocalGroupMember -Name $group.Name
    
     if($groupMember.length -le  0)
     {
         $localGroupMember =  [PSCustomObject]@{
                                    GroupName=($group.Name  | Out-String).Trim()
                                    User=""
                                    ObjectClass=""
                                    PrincipleSource  = ""
                                }
       
     }
     else
     {
            for($i = 0; $i -lt $groupMember.length; $i++){ 
               $localGroupMember =   [PSCustomObject]@{
                                        GroupName  = ($group.Name  | Out-String).Trim()
                                        User  = ($groupMember[$i].Name | Out-String).Trim()
                                        ObjectClass  = ($groupMember[$i].ObjectClass | Out-String).Trim()
                                        PrincipleSource  = ($groupMember[$i].PrincipalSource | Out-String).Trim()
                                        }
            

            }
      }
      $localGroupMember | Export-CSV $csvpath"Local-groupmember.csv" -Append -NoTypeInformation -Force
}
#Get-ChildItem -Recurse -Path $csvpath"*.csv"  | Get-Content | Add-Content $csvoutpath

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$wb = $excel.Workbooks.Add()

Get-ChildItem $csvpath\*.csv | ForEach-Object {
    if ((Import-Csv $_.FullName).Length -gt 0) {
        $csvBook = $excel.Workbooks.Open($_.FullName)
        $csvBook.ActiveSheet.Copy($wb.Worksheets($wb.Worksheets.Count))
        $csvBook.Close()
    }
}

