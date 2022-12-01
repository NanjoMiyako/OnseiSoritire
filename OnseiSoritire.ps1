$ps_speak=New-Object -ComObject SAPI.SpVoice

function SuffuleCards( $newList ){
    for($i=0; $i -lt ($newList.Count+20); $i++){
        $rnd = Get-Random -Minimum 0 -Maximum ($newList.Count-1)
        $tmp = $newList[$rnd]
        $newList[$rnd] = $newList[0]
        $newList[0] = $tmp
    }
}


function InitCards($newList){

    for($i=0; $i -lt 13; $i++){
        $str = "c" + ($i+1).ToString()
        [void]$newList.add($str)

        $str = "d" + ($i+1).ToString()
        [void]$newList.add($str)

        $str = "h" + ($i+1).ToString()
        [void]$newList.add($str)

        $str = "s" + ($i+1).ToString()
        [void]$newList.add($str)


    }
}

function OutputAllCards($headLists, $backLists, $dekki, $idx){
    for($i=0; $i -lt 7; $i++){
    write-host $headLists[$i]
    }
    write-host "aaa"
    for($i=0; $i -lt 7; $i++){
    write-host $backLists[$i]
    }
    write-host "bbb"

    $list1 = New-Object System.Collections.ArrayList
    $idx2 = $idx-1
    for($i=0; $i -lt $dekki.Count; $i++){
        if($idx2+1 -gt $dekki.Count-1){
            $idx2 = 0
        }else{
            $idx2 = $idx2 + 1
        }
        [void]$list1.add($dekki[$idx2])
    }
    write-host $list1

    write-host "ccc"
    write-host "d:" , $suits_d_heads
    write-host "s:" , $suits_s_heads
    write-host "c:" , $suits_c_heads
    write-host "h:" , $suits_h_heads
    write-host "ddd"

}

function InitPutCards($newList, $headLists, $backLists, $dekki){

    $idx1 = 0;
    for($i=0; $i -lt 7; $i++){
        $list1 = New-Object System.Collections.ArrayList
        [void]$list1.add($newList[$idx1])

        $headLists[$i] = $list1
        $idx1 = $idx1+1
        if($idx1 -eq 1){
            continue
        }

        $list2 = New-Object System.Collections.ArrayList
        for($j=0; $j -lt $i; $j++){
            [void]$list2.add($newList[$idx1])
            $idx1 = $idx1 + 1     
        }
        $backLists[$i] = $list2
    }

    for($k=$idx1; $k -lt $newList.Count; $k++){
        [void]$dekki.add($newList[$k])
    }
}

function IsCanMove([int]$pos1, [int]$pos2, $headLists, $backLists, $newList, [ref]$idx, $suits_d_heads, $suits_s_heads, $suits_c_heads, $suits_h_heads){

    if($pos2 -eq 8){
        return $false
    }elseif($pos1 -eq $pos2){
        return $false
    }elseif(($pos2 -gt 12) -or ($pos1 -gt 12)){
        return $false
    }elseif(($pos1 -lt 0) -or ($pos2 -lt 0)){
        return $false
    }

    if(($pos1 -ne 8) -and ($pos2 -lt 8)){
        if($headLists[$pos2].Count -gt 0){
            $head = $headLists[$pos2][$headLists[$pos2].Count-1]

            $num1 = $head.Substring(1)
            $suit1 = $head.Substring(0,1)

            for($i = $headLists[$pos1].Count-1; $i -ge 0; $i--){
                $cur = $headLists[$pos1][$i]

                $num2 = $cur.Substring(1)
                $suit2 = $cur.Substring(0,1)
                if([int]$num1 -eq ([int]$num2 + 1)){
                    if( ($suit1 -eq "h") -or ($suit1 -eq "d")){
                        if( ($suit2 -eq "c") -or ($suit2 -eq "s")){
                            return $true
                        }else{
                            return $false
                        }

                    }elseif( ($suit1 -eq "c") -or ($suit1 -eq "s")){
                        if( ($suit2 -eq "h") -or ($suit2 -eq "d")){
                            return $true
                        }
                        return $false
                    }
                }
            }
        }else{
            if($headLists[$pos1].Count -eq 0){
                return $false
            }
            $cur = $headLists[$pos1][0]

            $num2 = $cur.Substring(1)
            $suit2 = $cur.Substring(0,1)

            if($num2 -eq 13){
                return $true
            }else{
                return $false
            }

        }

        return $false
    }elseif(($pos1 -lt 7) -and ($pos2 -gt 8)){
       if($headLists[$pos1].Count -eq 0){
        return $false
       }
       $head = $headLists[$pos1][$headLists[$pos1].Count-1]
       $suit = $head.Substring(0,1)
       $num = [int]$head.Substring(1)

       if(($suit -eq "d") -and ($pos2 -eq 9)){
            if( ($suits_d_heads.Count -eq 0) -and ($num -eq 1)){
                return $true
            }else{
                $num2 = [int]$suits_d_heads[$suits_d_heads.Count-1].Substring(1)
                if($num2 -eq ($num-1)){
                    return $true
                }
                return $false
            }

       }elseif(($suit -eq "s") -and ($pos2 -eq 10)){
            if( ($suits_s_heads.Count -eq 0) -and ($num -eq 1)){
                return $true
            }else{
                $num2 = [int]$suits_s_heads[$suits_s_heads.Count-1].Substring(1)
                if($num2 -eq ($num-1)){
                    return $true
                }
                return $false
            }

       }elseif(($suit -eq "c") -and ($pos2 -eq 11)){
            if( ($suits_c_heads.Count -eq 0) -and ($num -eq 1)){
                return $true
            }else{
                $num2 = [int]$suits_c_heads[$suits_c_heads.Count-1].Substring(1)
                if($num2 -eq ($num-1)){
                    return $true
                }
                return $false
            }

       }elseif(($suit -eq "h") -and ($pos2 -eq 12)){
            if( ($suits_h_heads.Count -eq 0) -and ($num -eq 1)){
                return $true
            }else{
                $num2 = [int]$suits_h_heads[$suits_h_heads.Count-1].Substring(1)
                if($num2 -eq ($num-1)){
                    return $true
                }
                return $false
            }

       }
       return $false

    }elseif(($pos1 -eq 8) -and ($pos2 -lt 8)){
        $card1 = $newList[$idx.Value]
        $suit2 = $card1.Substring(0,1)
        $num2 = $card1.Substring(1)

        if($headLists[$pos2].Count -gt 0){
            $head = $headLists[$pos2][$headLists[$pos2].Count-1]

            $num1 = $head.Substring(1)
            $suit1 = $head.Substring(0,1)

            if([int]$num1 -eq ([int]$num2 +  1)){
                if( ($suit1 -eq "d") -or ($suit1 -eq "h") ){
                    if( ($suit2 -eq "s") -or ($suit2 -eq "c") ){
                        return $true
                    }
                }elseif( ($suit1 -eq "s") -or ($suit1 -eq "c") ){
                    if( ($suit2 -eq "d") -or ($suit2 -eq "h") ){
                        return $true
                    }
                }

                return $false

            }else{
                return $false
            }
        }else{
            if($num2 -eq 13){
                return $true
            }
            return $false
        }       
        
    }elseif( ($pos1 -eq 8) -and ($pos2 -gt 8) ){

        $card1 = $newList[$idx.Value]
        $suit2 = $card1.Substring(0,1)
        $num2 = $card1.Substring(1)

        if($pos2 -eq 9){
            if($suits_d_heads.Count -gt 0){
                $head = $suits_d_heads[$suits_d_heads.Count-1]

                $num1 = $head.Substring(1)
                $suit1 = $head.Substring(0,1)
            }else{
                $num1 = 0
                $suit1 = "d"
            }

        }elseif($pos2 -eq 10){
            if($suits_s_heads.Count -gt 0){
                $head = $suits_s_heads[$suits_s_heads.Count-1]

                $num1 = $head.Substring(1)
                $suit1 = $head.Substring(0,1)
            }else{
                $num1 = 0
                $suit1 = "s"
            }

        }elseif($pos2 -eq 11){
            if($suits_c_heads.Count -gt 0){
                $head = $suits_c_heads[$suits_c_heads.Count-1]

                $num1 = $head.Substring(1)
                $suit1 = $head.Substring(0,1)
            }else{
                $num1 = 0
                $suit1 = "c"
            }

        }elseif($pos2 -eq 12){
            if($suits_h_heads.Count -gt 0){
                $head = $suits_h_heads[$suits_h_heads.Count-1]

                $num1 = $head.Substring(1)
                $suit1 = $head.Substring(0,1)
            }else{
                $num1 = 0
                $suit1 = "h"
            }

        }

        if([int]$num1 -eq ([int]$num2 - 1)){
            if($suit1 -eq $suit2){
                return $true
            }
        }

        return $false
    }elseif( ($pos1 -gt 8) -and ($pos2 -lt 8)){
        if($pos1 -eq 9){
            if($suits_d_heads.Count -lt 1){
                return $false
            }
            $head = $suits_d_heads[$suits_d_heads.Count-1]
        }elseif($pos1 -eq 10){
            if($suits_s_heads.Count -lt 1){
                return $false
            }
            $head = $suits_s_heads[$suits_s_heads.Count-1]
        }elseif($pos1 -eq 11){
            if($suits_c_heads.Count -lt 1){
                return $false
            }
            $head = $suits_c_heads[$suits_c_heads.Count-1]
        }elseif($pos1 -eq 12){
            if($suits_h_heads.Count -lt 1){
                return $false
            }
            $head = $suits_h_heads[$suits_h_heads.Count-1]
        }
        $num1 = $head.Substring(1)
        $suit1 = $head.Substring(0,1)

        $head2 = $headLists[$pos2][$headLists[$pos2].Count-1]
        $suit2 = $head2.Substring(0,1)
        $num2 = $head2.Substring(1)

        if(([int]$num1+1) -eq [int]$num2){
            if( ($suit1 -eq "s") -or ($suit1 -eq "c") ){
                if( ($suit2 -eq "d") -or ($suit2 -eq "h") ){
                    return $true
                }
            }elseif( ($suit1 -eq "d") -or ($suit1 -eq "h") ){
                if( ($suit2 -eq "s") -or ($suit2 -eq "c") ){
                    return $true
                }
            }
        }

        return $false
    }

    return $false
    
}

function MoveCards([int]$pos1, [int]$pos2, $headLists, $backLists, $newList, [ref]$idx, $suits_d_heads, $suits_s_heads, $suits_c_heads, $suits_h_heads){

    if($pos2 -eq 8){
        return
    }elseif($pos1 -eq $pos2){
        return
    }elseif(($pos2 -gt 12) -or ($pos1 -gt 12)){
        return
    }elseif(($pos1 -lt 0) -or ($pos2 -lt 0)){
        return
    }

    if(($pos1 -ne 8) -and ($pos2 -lt 7)){

        if($headLists[$pos2].Count -gt 0){
            $head = $headLists[$pos2][$headLists[$pos2].Count-1]

            $num1 = $head.Substring(1)
            $suit1 = $head.Substring(0,1)

            for($i = $headLists[$pos1].Count-1; $i -ge 0; $i--){
                $cur = $headLists[$pos1][$i]

                $num2 = $cur.Substring(1)
                $suit2 = $cur.Substring(0,1)
                if([int]$num1 -eq ([int]$num2 + 1)){
                    if( ($suit1 -eq "h") -or ($suit1 -eq "d")){
                        if( ($suit2 -eq "c") -or ($suit2 -eq "s")){
                            $mov_st = $i
                            for($j = $mov_st; $j -lt $headLists[$pos1].Count; $j++){
                                [void]$headLists[$pos2].add($headLists[$pos1][$j])   
                            }
                            $list1 = New-Object System.Collections.ArrayList
                            for($k=0;  $k -lt $mov_st; $k++){
                                [void]$list1.add($headLists[$pos1][$k])
                            }
                            $headLists[$pos1] = $list1
                            if($headLists[$pos1].Count -eq 0){
                                if($backLists[$pos1].Count -gt 0){
                                    $tail = $backLists[$pos1][$backLists[$pos1].Count-1]
                                    [void]$headLists[$pos1].add($tail)
                                    $backLists[$pos1].remove($tail)
                                }
                            }
                            return
                        }else{
                            return
                        }
                    }elseif( ($suit1 -eq "c") -or ($suit1 -eq "s")){
                        if( ($suit2 -eq "h") -or ($suit2 -eq "d")){
                            $mov_st = $i
                            for($j = $mov_st; $j -lt $headLists[$pos1].Count; $j++){
                                [void]$headLists[$pos2].add($headLists[$pos1][$j])   
                            }
                            $list1 = New-Object System.Collections.ArrayList
                            for($k=0;  $k -lt $mov_st; $k++){
                                [void]$list1.add($headLists[$pos1][$k])
                            }
                            $headLists[$pos1] = $list1
                            if($headLists[$pos1].Count -eq 0){
                                if($backLists[$pos1].Count -gt 0){
                                    $tail = $backLists[$pos1][$backLists[$pos1].Count-1]
                                    [void]$headLists[$pos1].add($tail)
                                    $backLists[$pos1].remove($tail)
                                }
                            }

                            return
                        }
                        return
                    }
                }
            }
        }else{

            $headLists[$pos2] = $headLists[$pos1]
            $list1 = New-Object System.Collections.ArrayList
            $headLists[$pos1] = $list1
            if($backLists[$pos1].Count -gt 0){
                $bh = $backLists[$pos1][$backLists[$pos1].Count-1]
                [void]$headLists[$pos1].add($bh)
                $backLists[$pos1].remove($bh)
            }

        }

        return

    }elseif(($pos1 -lt 7) -and ($pos2 -gt 8)){

       $head = $headLists[$pos1][$headLists[$pos1].Count-1]
       $suit = $head.Substring(0,1)
       $num = [int]$head.Substring(1)

       if(($suit -eq "d") -and ($pos2 -eq 9)){
            if( ($suits_d_heads.Count -eq 0) -and ($num -eq 1)){
                [void]$suits_d_heads.add($head)
                $headLists[$pos1].remove($head)

            }else{
                [void]$suits_d_heads.add($head)
                $headLists[$pos1].remove($head)

            }
            

       }elseif(($suit -eq "s") -and ($pos2 -eq 10)){
            if( ($suits_s_heads.Count -eq 0) -and ($num -eq 1)){
                [void]$suits_s_heads.add($head)
                $headLists[$pos1].remove($head)

            }else{
                [void]$suits_s_heads.add($head)
                $headLists[$pos1].remove($head)
            }

       }elseif(($suit -eq "c") -and ($pos2 -eq 11)){
            if( ($suits_c_heads.Count -eq 0) -and ($num -eq 1)){
                [void]$suits_c_heads.add($head)
                $headLists[$pos1].remove($head)

            }else{
                [void]$suits_c_heads.add($head)
                $headLists[$pos1].remove($head)

            }

       }elseif(($suit -eq "h") -and ($pos2 -eq 12)){
            if( ($suits_h_heads.Count -eq 0) -and ($num -eq 1)){
                [void]$suits_h_heads.add($head)
                $headLists[$pos1].remove($head)

            }else{
                [void]$suits_h_heads.add($head)
                $headLists[$pos1].remove($head)

            }

       }

       if($headLists[$pos1].Count -eq 0){
            if($backLists[$pos1].Count -gt 0){
                $tail = $backLists[$pos1][$backLists[$pos1].Count-1]
                [void]$headLists[$pos1].add($tail)
                $backLists[$pos1].Remove($tail)
            }
       }
       
    }elseif(($pos1 -eq 8) -and ($pos2 -lt 7)){
        $card1 = $newList[$idx.Value]
        $suit2 = $card1.Substring(0,1)
        $num2 = $card1.Substring(1)

        if($headLists[$pos2].Count -gt 0){
            $head = $headLists[$pos2][$headLists[$pos2].Count-1]

            $num1 = $head.Substring(1)
            $suit1 = $head.Substring(0,1)

            if([int]$num1 -eq ([int]$num2 +  1)){
                if( ($suit1 -eq "d") -or ($suit1 -eq "h") ){
                    if( ($suit2 -eq "s") -or ($suit2 -eq "c") ){
                        $headLists[$pos2].add($card1)
                        $newList.remove($card1)
                        $idx.Value = $idx.Value-1

                        return
                    }
                }elseif( ($suit1 -eq "s") -or ($suit1 -eq "c") ){
                    if( ($suit2 -eq "d") -or ($suit2 -eq "h") ){
                        $headLists[$pos2].add($card1)
                        $newList.remove($card1)
                        $idx.Value = $idx.Value-1

                        return
                    }
                }

                return

            }else{
                return
            }
        }else{
            if($num2 -eq 13){
                $headLists[$pos2].add($card1)
                $newList.remove($card1)
                $idx.Value = $idx.Value-1
                return
            }
            return
        }       
        
    }elseif( ($pos1 -eq 8) -and ($pos2 -gt 8) ){

        $card1 = $newList[$idx.Value]
        $suit2 = $card1.Substring(0,1)
        $num2 = $card1.Substring(1)

        if($pos2 -eq 9){
            if($suits_d_heads.Count -gt 0){
                $head = $suits_d_heads[$suits_d_heads.Count-1]

                $num1 = $head.Substring(1)
                $suit1 = $head.Substring(0,1)
            }else{
                $num1 = 0
                $suit1 = "d"
            }

        }elseif($pos2 -eq 10){
            if($suits_s_heads.Count -gt 0){
                $head = $suits_s_heads[$suits_s_heads.Count-1]

                $num1 = $head.Substring(1)
                $suit1 = $head.Substring(0,1)
            }else{
                $num1 = 0
                $suit1 = "s"
            }

        }elseif($pos2 -eq 11){
            if($suits_c_heads.Count -gt 0){
                $head = $suits_c_heasd[$suits_c_heads.Count-1]

                $num1 = $head.Substring(1)
                $suit1 = $head.Substring(0,1)
            }else{
                $num1 = 0
                $suit1 = "c"
            }

        }elseif($pos2 -eq 12){
            if($suits_h_heads.Count -gt 0){
                $head = $suits_h_heads[$suits_h_heads.Count-1]

                $num1 = $head.Substring(1)
                $suit1 = $head.Substring(0,1)
            }else{
                $num1 = 0
                $suit1 = "h"
            }

        }

        if([int]$num1 -eq ([int]$num2 - 1)){
            if($suit1 -eq $suit2){
                if($suit1 -eq "d"){
                    $suits_d_heads.add($card1)                    
                }elseif($suit1 -eq "s"){
                    $suits_s_heads.add($card1)   
                }elseif($suit1 -eq "c"){
                    $suits_c_heads.add($card1)   
                }elseif($suit1 -eq "h"){
                    $suits_h_heads.add($card1)   
                }
                $newList.remove($card1)
                $idx.Value = $idx.Value-1
                return
            }
        }

        return 
    }elseif( ($pos1 -gt 8) -and ($pos2 -lt 7)){
        $head = ""
        if($pos1 -eq 9){
            if($suits_d_heads.Count -lt 1){
                return
            }
            $head = $suits_d_heads[$suits_d_heads.Count-1]
        }elseif($pos1 -eq 10){
            if($suits_s_heads.Count -lt 1){
                return
            }
            $head = $suits_s_heads[$suits_s_heads.Count-1]
        }elseif($pos1 -eq 11){
            if($suits_c_heads.Count -lt 1){
                return
            }
            $head = $suits_c_heads[$suits_c_heads.Count-1]
        }elseif($pos1 -eq 12){
            if($suits_h_heads.Count -lt 1){
                return
            }
            $head = $suits_h_heads[$suits_h_heads.Count-1]
        }
        $num1 = $head.Substring(1)
        $suit1 = $head.Substring(0,1)

        $head2 = $headLists[$pos2][$headLists[$pos2].Count-1]
        $suit2 = $head2.Substring(0,1)
        $num2 = $head2.Substring(1)

        if(([int]$num1+1) -eq [int]$num2){
            if( ($suit1 -eq "s") -or ($suit1 -eq "c") ){
                if( ($suit2 -eq "d") -or ($suit2 -eq "h") ){

                    $headLists[$pos2].add($head)
                    if($suit1 -eq "s"){
                        $suits_s_heads.remove($head)
                    }else{
                        $suits_c_heads.remove($head)
                    }
                    return
                }
            }elseif( ($suit1 -eq "d") -or ($suit1 -eq "h") ){
                if( ($suit2 -eq "s") -or ($suit2 -eq "c") ){

                    $headLists[$pos2].add($head)
                    if($suit1 -eq "d"){
                        $suits_d_heads.remove($head)
                    }else{
                        $suits_h_heads.remove($head)
                    }
                    return $true
                }
            }
        }

        return $false
    }

}

function GetFrontHeadAndFrontEndCards($pos1, $headLists, $newList, $idx, $suits_d_heads, $suits_s_heads, $suits_c_heads, $suits_h_heads){
    
    $retVal = New-Object System.Collections.ArrayList
    if($pos1 -lt 7){
        if($headLists[$pos1].Count -gt 0){
            $head = $headLists[$pos1][$headLists[$pos1].Count-1]
            $end = $headLists[$pos1][0]

            [void]$retVal.add($head)
            [void]$retVal.add($end)

        }
    }elseif($pos1 -eq 8){
        $retVal.add($newList[$idx])
    }elseif( ($pos1 -gt 8) -and ($pos1 -lt 13) ){
        if($pos1 -eq 9){
            $retVal.add($suits_d_heads[$suits_d_heads.Count-1])
        }elseif($pos1 -eq 10){
            $retVal.add($suits_s_heads[$suits_s_heads.Count-1])
        }elseif($pos1 -eq 11){
            $retVal.add($suits_c_heads[$suits_c_heads.Count-1])
        }elseif($pos1 -eq 12){
            $retVal.add($suits_h_heads[$suits_h_heads.Count-1])
        }
    }

    return $retVal
}


function DrawDekki($newList, [ref]$idx){
    if($idx.Value -eq ($newList.Count-1)){
        $idx.Value = 0
    }else{
        $idx.Value = $idx.Value + 1
    }

    $retVal = $newList[$idx.Value]
    return $retVal
}

function GetBackCardCount($pos1){
    if(($pos1 -gt 0) -and ($pos1 -lt 7)){
        return $backLists[$pos1].Count
    }
    return 
}

function IsClear($backLists){
    for($i=0; $i -lt 7; $i++){
        if($backLists[$i].Count -gt 0){
            return $false
        }
    }

    return $true
}
function GetSuitStr($suit){
    if($suit -eq "d"){
        return "ダイヤ"
    }elseif($suit -eq "s"){
        return "スペード"
    }elseif($suit -eq "c"){
        return "クラブ"
    }elseif($suit -eq "h"){
        return "ハート"
    }
    return ""    
}

function GetNumberStr($num){

    if($num -eq 13){
        return "キング"
    }elseif($num -eq 12){
        return "クイーン"
    }elseif($num -eq 11){
        return "ジャック"
    }
    return $num
}


function ExeCmd($cmdStr){


    if( ($cmdStr.IndexOf("-") -ne -1) -and ($cmdStr.IndexOf("*") -eq -1)){
        
        $pairs = $cmdstr.split("-")
        if([int]$pairs[0] -lt 8){
            $pos1 = [int]$pairs[0]-1
        }else{
            $pos1 = [int]$pairs[0]
        }
        if([int]$pairs[1] -lt 8){
            $pos2 = [int]$pairs[1]-1
        }else{
            $pos2 = [int]$pairs[1]
        }

        $ret1 = IsCanMove $pos1 $pos2 $headLists $backLists $dekki ([ref]$idx) $suits_d_heads $suits_s_heads $suits_c_heads $suits_h_heads

        if($ret1 -eq $false){
            $str2 = "無効な移動です"
            Write-host $str2
            $ps_speak.Speak($str2)

        }else{
     　　　MoveCards $pos1 $pos2 $headLists $backLists $dekki ([ref]$idx) $suits_d_heads $suits_s_heads $suits_c_heads $suits_h_heads
           $str = $pairs[0] + "番のカードを" + $pairs[1] + "番に移動しました。"
           Write-host $str
           $ps_speak.Speak($str)

        }


        $ret2 = IsClear $backLists
        if($ret2 -eq $true){
            $str2 = "ゲームクリアです"
            Write-host $str2
            $ps_speak.Speak($str2)

            exit
        }

    }elseif($cmdStr -eq "+"){
       $str2 = "デッキから一枚引きます"
       Write-host $str2
       $ps_speak.Speak($str2)

       $ret = DrawDekki $dekki ([ref]$idx)
       $suit = $ret.Substring(0,1)
       $num = $ret.Substring(1)
       
       $str = "デッキのカードは"
       $str += GetSuitStr $suit
       $str += "の"
       $str += GetNumberStr $num
       $str += "です。"
       Write-host $str
       $ps_speak.Speak($str)
       
    }elseif($cmdStr.IndexOf("**") -ne -1){
        $num3 = [int]($cmdStr.Substring(2))
        if($num3 -lt 8){
            $pos1 = ([int]$num3-1)
        }else{
            $pos1 = $num3
        }


        if($pos1 -lt 7){

            $str = [string]$num3
            $str += "番のカードは一番上から"
            for($i=0; $i -lt $headLists[$pos1].Count; $i++){

               $v1 = $headLists[$pos1][$i]
               $suit = $v1.Substring(0,1)
               $num = $v1.Substring(1)

               $str += "、"
               $str += GetSuitStr $suit
               $str += "の"
               $str += GetNumberStr $num
            
            }
            $str += "、です。"
            Write-host $str
            $ps_speak.Speak($str)

        }elseif(($pos1 -ge 9) -and ($pos1 -le 12)){ 

            $str = [string]$num3
            $str += "番の先頭のカードは、"        
            if($pos1 -eq 9){
                if($suits_d_heads.Count -eq 0){
                    $str += "ありません。"
                    write-host $str
                    $ps_speak.Speak($str)

                    return
                }
                $v1 = $suits_d_heads[$suits_d_heads.Count-1]
            }elseif($pos1 -eq 10){
                if($suits_s_heads.Count -eq 0){
                    $str += "ありません。"
                    write-host $str
                    $ps_speak.Speak($str)

                    return
                }
                $v1 = $suits_s_heads[$suits_s_heads.Count-1]
            }elseif($pos1 -eq 11){
                if($suits_c_heads.Count -eq 0){
                    $str += "ありません。"
                    write-host $str
                    $ps_speak.Speak($str)

                    return
                }
                $v1 = $suits_c_heads[$suits_c_heads.Count-1]
            }elseif($pos1 -eq 12){
                if($suits_h_heads.Count -eq 0){
                    $str += "ありません。"
                    write-host $str
                    $ps_speak.Speak($str)

                    return
                }
                $v1 = $suits_h_heads[$suits_h_heads.Count-1]
            }
            $suit = $v1.Substring(0,1)
            $num = $v1.Substring(1)

            $str += "、"
            $str += GetSuitStr $suit
            $str += "の"
            $str += GetNumberStr $num
            $str += "、です。"

            write-host $str
            $ps_speak.Speak($str)

            return
        }

    }elseif(($cmdStr.IndexOf("*") -ne -1) -and ($cmdStr.IndexOf("-") -eq -1)){
        $num2 = [int]$cmdStr.Substring(1)
        if($num2 -lt 8){
            $pos1 = [int]$num2-1
        }else{
            $pos1 = [int]$num2
        }
        $ret = GetFrontHeadAndFrontEndCards $pos1 $headLists $newList $idx $suits_d_heads $suits_s_heads $suits_c_heads $suits_h_heads

        $str = ""
        if($ret.Count -ge 1){
            $v1 = $ret[0]
            $suit = $v1.Substring(0,1)
            $num = $v1.Substring(1)

            $str = [string]$num2
            $str += "番のカードの一番上のカードは"
            $str += "、"
            $str += GetSuitStr $suit
            $str += "の"
            $str += GetNumberStr $num
            $str += "です。"

            Write-host $str
            $ps_speak.Speak($str)

        }

        if($ret.Count -ge 2){
            $v2 = $ret[1]
            $suit = $v2.Substring(0,1)
            $num = $v2.Substring(1)

            $str = [string]$num2
            $str += "番のカードの一番下のカードは"
            $str += "、"
            $str += GetSuitStr $suit
            $str += "の"
            $str += GetNumberStr $num
            $str += "です。"

            Write-host $str
            $ps_speak.Speak($str)

        }
    }elseif($cmdStr.IndexOf("*-") -ne -1){
        $p1 = $cmdStr.Substring(2)
        if($p1 -lt 8){
            $pos1 = $p1-1
        }
        $restCount = GetBackCardCount $pos1

        $str = $p1
        $str += "番の残りカードの個数は、"
        $str += $restCount
        $str += "枚です。"

        Write-host $str
        $ps_speak.Speak($str)

        return

    }elseif($cmdStr -eq "++"){
        $str2 = "ゲームを終了します"
        Write-host $str2
        $ps_speak.Speak($str2)
        exit
    }

    return
}




$cards = New-Object System.Collections.ArrayList
InitCards $cards
SuffuleCards $cards

<# デバッグ用

$cards = @( 'h8', 'd1', 'd11', 's4', 'h13', 's2', 'c5',
'h1', 
'h10', 's1',
'c2', 'h3', 'h4',
'd10', 'c4', 'c12', 'd3'
'c11', 'd5', 'h5', 's5', 'c6',
'h6', 's6', 'd9', 'h7', 'd8', 'h2'
'd13','c13', 'd7', 'd4', 'c7', 'd6', 'h9', 'c1', 'c10', 's8', 's12', 'c3', 'c9', 's10', 's11', 'h11', 'd12', 'd2', 's7', 's3', 's13', 'c8','h12')

#>
$headLists = New-Object System.Collections.ArrayList
$backLists = New-Object System.Collections.ArrayList

$dekki= New-Object System.Collections.ArrayList
$idx = 0

$suits_d_heads = New-Object System.Collections.ArrayList
$suits_h_heads = New-Object System.Collections.ArrayList
$suits_c_heads = New-Object System.Collections.ArrayList
$suits_s_heads = New-Object System.Collections.ArrayList


for($i=0; $i -lt 7; $i++){
 $list1 = New-Object System.Collections.ArrayList
 [void]$headLists.add($list1)

 $list2 = New-Object System.Collections.ArrayList
 [void]$backLists.add($list2)

}
InitPutCards $cards $headLists $backLists $dekki

while(1){

    $str2 = "コマンドを入力してください"
    $ps_speak.Speak($str2)
    $cmd = Read-Host $str2

    $cmd2 = $cmd.Replace("*", "、あすたりすく、")
    $cmd2 = $cmd.Replace("-", "、はいふん、")
    $str2 = "入力したコマンド、"+$cmd2
    $ps_speak.Speak($str2)

    ExeCmd $cmd
    
    ## デバッグ用
    ## write-host "---"
    ## OutputAllCards $headLists $backLists $dekki $idx

}
