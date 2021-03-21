##https://devblogs.microsoft.com/scripting/speed-up-array-comparisons-in-powershell-with-a-runtime-regex/

$a = "red.","blue.","yellow.","green.","orange.","purple."
$b = "blue.","green.","orange.","white.","gray."
"a:$a"
"b:$b"

$b |? {$a -contains $_}
"blue.green.orange"

$b |? {$a -notcontains $_}
"white.gray."

$a |foreach {[regex]::escape($_)}

$a_regex = "^(red\.|blue\.|yellow\.|green\.|orange\.|purple\.)$"
$b -match $a_regex
$b -notmatch $a_regex

'(?i)^(' + (($a |foreach {[regex]::escape($_)}) –join "|") + ')$'
(?i)^(red\.|blue\.|yellow\.|green\.|orange\.|purple\.)$

[regex] $a_regex = '(?i)^(' + (($a |foreach {[regex]::escape($_)}) –join "|") + ')$'
$a_regex.tostring()
"(?i)^(red\.|blue\.|yellow\.|green\.|orange\.|purple\.)$"


[regex] $a_regex = '(?i)^(' + (($a |foreach {[regex]::escape($_)}) –join "|") + ')$'
$b -match $a_regex
$b -notmatch $a_regex
