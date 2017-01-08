param($length =44)

$ascii=$NULL;
For ($a=48;$a -le 122;$a++) {$ascii+=,[char][byte]$a }
For ($loop=1; $loop -le $length; $loop++) {
    $TempPassword+=($ascii | GET-RANDOM)
}
return $TempPassword
