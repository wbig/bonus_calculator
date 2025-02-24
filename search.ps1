function regex-from-lines($pattern, $listfile){
    foreach($i in $listfile){
        $content = Get-Content -Path $i -Encoding UTF8
        $row = 0
        $t = 0
        foreach($line in $content){
            $row ++
            $res = $line | Select-String -Pattern $pattern
            if ($res){
                $t ++
                if($t -eq 1){ Write-Host $i }
                Write-Host $row $res
            }
        }
    }
}

cls
$listfile = Get-ChildItem "*.py"
# $listfile = "summarizing.py"

$pattern = "Êµ·¢½ð¶î"

regex-from-lines $pattern $listfile