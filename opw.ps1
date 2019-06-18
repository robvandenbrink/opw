$nodenets = @()
$domainmembers = get-adcomputer -filter *
foreach ($node in $domainmembers) {
    $netlist = iex ("./psexec /accepteula \\"+$node.name +" netsh wlan show profiles") 2>./a | Select-String -Pattern ": "
    if(($netlist -like "*was not found*") -or ($netlist.length -eq 0)) { write-host "No Wireless on host " $node.name }
    else {
      write-host "Assessing Wireless on host " $node.name
      foreach ($net in $netlist) {
        [console]::write(".")
        $netprf = ($net -split(": "))[1]
        $cmd = "./psexec /accepteula \\"+$node.name +" netsh wlan show profiles name="+ "`'"+$netprf+"`'"
        $netparmlist = iex $cmd 2>./a
        $netparmlist2 = $netparmlist | select-string -pattern ": " | select-string -pattern "Applied" -NotMatch | select-string -pattern "Profile" -NotMatch
        $x = New-Object psobject
        $x | add-member -membertype NoteProperty -name "Node" -Value $node.name
        foreach($parm in $netparmlist2) {
          $t1 = $parm -split ": "
          $x | add-member –membertype NoteProperty –name ($t1[0].trim(" ")) –Value ($t1[1]) ;
          }
        $nodenets += $x
        }
      }
  }
$nodenets | select Node, Name, "Connection Mode", "SSID Name", Authentication, Cipher, "Security Key" | Out-GridView
