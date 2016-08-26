    #####################################################################################################################################################################
    # Graphs and HTML report
    #####################################################################################################################################################################
Function Select-RandomMove {$movelist[(get-random -maximum ([array]$movelist).count)]}

$HTMLFile = "C:\temp\generatemove.html"
[Array]$housemovelist =@(
"African Step",
"Scissors",
"Kerry Step (Cuban Salsa 4 point heel toe)",
"Loose Legs",
"Heel Toe Hop",
"Cross Step Drag",
"Pas De Bourree",
"Pivoting Pas De Bourree",
"Gallop (Slow)",
"Gallop (Fast)",
"Skate",
"Train",
"Farmer",
"Swirl",
"Pas De Bourree Loop",
"Pow Wow"
"Set-Up",
"Salsa Hop",
"Shuffle",
"Jack in the Box",
"Scribble Foot",
"Crossroads",
"SpongeBob",
"Heel Step Variation",
"Pivot Groove Step",
"Sidewalk"
"Crosswalk",
"Happy Feet",
"Cross Step x 2",
"Heel Step",
"Kriss Kross",
"Around the World",
"Stomp",
"Salsa Step",
"Roger Rabbit - Reject",
"Pivoting Skate"
)

$timeout = new-timespan -Minutes 5 # Length of time for entire dance session
$intervaltiming = 10 # Number of seconds between moves


$sw = [diagnostics.stopwatch]::StartNew()
$movelist = $housemovelist
$move = Select-RandomMove
$move
# Create an HTML Page that lists each move

$html = @"
<!DOCTYPE html>
<meta http-equiv="refresh" content="3">
<html>
<body>
<div style="position:relative">
<center><p style="font-size:150px">$move</p></center>
</div>
</body>
</html>
"@
#>

ConvertTo-Html -body $html | Out-File $HTMLFile

$ie = New-Object -ComObject InternetExplorer.Application
$ie.Visible = $true
$asm = [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
$screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
$ie.height = $screen.height
$ie.width = $screen.width
$ie.Navigate($HTMLFile)

Start-Sleep -Seconds $intervaltiming
while ($sw.elapsed -lt $timeout){
#while($true){
$lastmove = $move
while($lastmove -like $move){$move = Select-RandomMove}

$html = @"
<!DOCTYPE html>
<meta http-equiv="refresh" content="3">
<html>
<body>
<div style="position:relative">
<center><p style="font-size:150px">$move</p></center>
</div>
</body>
</html>
"@
$move
ConvertTo-Html -body $html | Out-File $HTMLFile
Start-Sleep -Seconds $intervaltiming
#}
}

write-host "Timed out"
$move = "Finished"
$html = @"
<!DOCTYPE html>
<meta http-equiv="refresh" content="3">
<html>
<body>
<div style="position:relative">
<center><p style="font-size:150px">$move</p></center>
</div>
</body>
</html>
"@
ConvertTo-Html -body $html | Out-File $HTMLFile