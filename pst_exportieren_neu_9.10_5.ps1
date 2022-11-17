############################################################################################################################
#					Script to export Mailboxes to PST
############################################################################################################################

#Mit dem Cmdlet Import-CSV wird eine CSV-Datei importiert und Informationen 
#in einer Variable $Mailboxes gespeichert. 


#$Mailboxes1 = import-csv -path   "\\Server1\d$\export_Postfach\pst_liste1.csv" -delimiter ';'

#$cred = Get-Credential
 
#$TargetPst -Variable enthält ein Share-Verzeichnis auf Zielserver
$TargetPst =  "\\Server1\d$\export_Postfach\"


# Liste der fehlgeschlagenen Mailboxen
# $FailedMailboxes=@()

$FailedMailboxes = $TargetPst +"pst_fail.txt"
$pstlog = $TargetPst + "pst_log.txt"

# Begrenzt die Anzahl der zu verarbeitenden Mailboxen
#$Mailboxes1 = $Mailboxes

#Write-Host $Mailboxes1 [4]


# Mit der For-Each-Schleife werden $AdUsers Objekte iteriert (eine Liste schrittweise durchgehen). Mit dem Cmdlet Get-AdUser 
# wird mithilfe von SAMAccountName ein Active Directory-Benutzerobjekt abgerufen und ein 
# Anzeigenbenutzerobjekt in der $ad Variablen gespeichert.
	
# jede einzelne Zeile aus $Mailboxes wird an $Mailbox übergeben	
	ForEach ($Mailbox in $Mailboxes1){
    
	Write-Host $Mailbox
	# Alias/mailNickname werden aus Liste in $Mailbox übergeben	
	$alias = $Mailbox.alias
	$z = $Mailbox.mail
    
    	
# mit Get-ADUser wird das Directory durchsucht und mit Alias verglichen
    $ad =Get-ADUser -filter 'mailNickname -eq $alias'
    

# Anweisung was mit $ad geschehen soll  
	If ($ad -eq '$null'){
		Write-Host -ForegroundColor Red "Processing of Mailbox $alias failed!" 
        Write-Host -ForegroundColor DarkGreen `n
		#new-item . -name "pst_fail.txt" -value $z -force
        #$FailedMailboxes  | set-content $FailedMailboxes 
        $z | Add-content $FailedMailboxes
        get-content $FailedMailboxes
    }    
	else {
# Postfach wird auf das angegebene Ziel als .pst abgelegt	           
        Write-Host -ForegroundColor DarkGreen "Starting Export for Mailbox $z..."
        Write-Host -ForegroundColor DarkGreen `n 
        $nameexp = "Export-" +$z
        $FilePathExp = $TargetPst+$alias+".pst" 
        #Get-Mailbox $z | New-MailboxExportRequest -Name $nameexp -FilePath $FilePathExp        
        #New-MailboxExportRequest -Name $alias -Mailbox $z -FilePath $FilePathExp 
         
            
        #Write-Host -ForegroundColor DarkGreen `n
	
# Mit nachfolgendem Befehl kann der Status aller Aufträge mit Fortschritt in Prozent ausgeben werden. 
	    Get-MailboxExportRequest | Get-MailboxExportRequestStatistics | Format-Table -AutoSize
    
    }

    # Schleife die wartet bis der Export abgeschlossen ist
        #$RequestStatus = Get-MailboxExportRequest -Name "$alias"
        #Write-Host "$RequestStatus"         
        #New-MailboxExportRequest -Name $alias -Mailbox $z -FilePath $FilePathExp #  -AcceptLargeDataLoss -ErrorAction Stop
    #while ($RequestStatus.Status -ne "Completed") {
        ##Get-MailboxExportRequest | Remove-MailboxExportRequest -Confirm:$false
        #Write-Host -ForegroundColor DarkGreen "Export for Mailbox $alias still running..."
        #$RequestStatus = Get-MailboxExportRequest -Name $alias ; Start-Sleep 10

        #Get-Mailbox $z | New-MailboxExportRequest -Name $nameexp  -FilePath $FilePathExp -BadItemLimit unlimited -AcceptLargeDataLoss -ErrorAction Stop
        #New-MailboxExportRequest -Name $alias -Mailbox $alias -FilePath $FilePathExp -BadItemLimit unlimited -AcceptLargeDataLoss -ErrorAction Stop
        #New-MailboxExportRequest -Name $alias -Mailbox $z -FilePath $FilePathExp #  -AcceptLargeDataLoss -ErrorAction Stop
        Get-MailboxExportRequest | Get-MailboxExportRequestStatistics -IncludeReport | fl  >> $pstlog 
#}
    # Sofern nicht anders angegeben, wird die Fehleraktionseinstellung auf den Wert stop festgelegt, 
    # und die Zeile $ErrorActionPreference = 'stop'

    try {
        $ErrorActionPreference = "Stop";
        
        Write-Host -ForegroundColor DarkGreen "Starting Export for Mailbox $alias..."
        Write-Host -ForegroundColor DarkGreen `n

    # Der Parameter -BadItemLimit gibt die maximale Anzahl der fehlerhaften Elemente an, die zulässig sind, 
    # bevor die Anforderung fehlschlägt. Ein fehlerhaftes Element ist ein beschädigtes Element im Quellpostfach, 
    # das nicht in das Zielpostfach kopiert werden kann.

    # Der Schalter -AcceptLargeDataLoss legt fest, dass die Anforderung auch dann fortgesetzt werden soll, 
    # wenn eine große Anzahl von Elementen im Quellpostfach nicht in das Zielpostfach kopiert werden kann.

    # Wenn die Ausführung des PowerShell-Skripts anhalten soll, wenn bei einem Aufruf von Stop-Process ein Fehler auftritt, 
    # kann einfach der Parameter -ErrorAction hinzugefügt und der Wert Stop verwendet werden

        New-MailboxExportRequest -Name $alias -Mailbox $z -FilePath $FilePathExp -ErrorAction Continue
        #Get-Mailbox $Mailbox | New-MailboxExportRequest -Name $nameexp -FilePath $FilePathExp -BadItemLimit unlimited -acceptdatalost -AcceptLargeDataLoss -ErrorAction Stop
        
        Write-Host -ForegroundColor DarkGreen `n

    }
    # Hier im catch findet die Fehlerbehandlung statt, z.B. das Schreiben eines Logs
	# Der letzte aufgezeichnete Fehler ist hier über die Variable $_ abrufbar,
    # einzelne Eigenschaften daher nach diesem Muster: $_.Exception.Message
    catch {

        # Ups...
        Write-Host -ForegroundColor Red "Processing of Mailbox $alias failed!"
        Write-Host -ForegroundColor DarkGreen `n

        # Namen merken
        # $FailedMailboxes += New-Object -TypeName psobject -Property @{Mailbox="$z"} 
                  
        # Nächste Mailbox
        continue

    }
        <#
        Jede Anweisung in diesem Block finally wird immer ausgeführt, egal ob ein
        Fehler aufgetreten ist oder nicht. Dieser Block ist optional.
        #>
    finally{

        # AufrÃ¤umen am Ende 
        if ( $RequestStatus.Status -eq "Completed" ) {
        
            Write-Host -ForegroundColor DarkGreen "Export for Mailbox $alias done. Cleaning up..."
        
        # Nachdem die Daten exportiert wurden, kann nun der Export Request gelöscht werden. 
        # Um alle Requests zu löschen, muss folgender Befehl verwendet werden:
            
            #Get-MailboxExportRequest | Remove-MailboxExportRequest -Confirm:$false
            #Remove-PSSession $session
        }   
       } 
    }



# Fehlgeschlagene Mailboxen ausgeben
    Write-Host -ForegroundColor DarkGreen "Unexported mailboxes:"
    Write-Host -ForegroundColor DarkGreen `n
    $FailedMailboxes | Format-Table -AutoSize >> $TargetPst\pst_fail.txt


