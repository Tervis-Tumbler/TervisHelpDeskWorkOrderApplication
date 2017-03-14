#Requires -Modules KanbanizePowerShell, TrackITUnOfficial, TrackITWebAPIPowerShell, get-MultipleChoiceQuestionAnswered, TervisTrackITWebAPIPowerShell, MailToURI
#Requires -Version 4
#Requires -RunAsAdministrator

Function Install-TervisHelpDeskWorkOrderApplication {
    fsutil behavior set SymlinkEvaluation L2L:1 R2R:1 L2R:1 R2L:1
    if (-not (Get-KanbanizeAPIKey)) {Install-TervisKanbanize}
}

function Invoke-PrioritizeConfirmTypeAndMoveCard {
    [CmdletBinding()]
    param()

    $HelpDeskProcessBoardID = Get-TervisKanbanizeHelpDeskBoardIDs -HelpDeskProcess
    $HelpDeskTechnicianProcessBoardID = Get-TervisKanbanizeHelpDeskBoardIDs -HelpDeskTechnicianProcess

    $Types = get-TervisKanbanizeTypes

    $WaitingToBePrioritized = Get-KanbanizeTervisHelpDeskCards -HelpDeskTriageProcess |
    where columnpath -NotMatch "Waiting for scheduled date" |
    sort positionint

    $global:CardsThatNeedToBeCreatedTypes = @()
    $global:ToBeCreatedTypes = @()

    foreach ($Card in $WaitingToBePrioritized) {
        Get-TervisWorkOrderDetails -Card $Card

        read-host "Hit enter once you have reviewed the details about this request"

        $SkipCard = get-MultipleChoiceQuestionAnswered -Question "Skip this card and move to the next one?" -Choices "Yes","No" | ConvertTo-Boolean               
        if ($SkipCard) { continue }

        if ($Card.Type -ne "None") {
            $TypeCorrect = get-MultipleChoiceQuestionAnswered -Question "Type ($($Card.Type)) correct?" -Choices "Yes","No" | ConvertTo-Boolean               
        }
        
        if (-not $TypeCorrect -or ($Card.Type -eq "None")) {        
            $SelectedType = $Types | Out-GridView -PassThru
    
            if ($SelectedType -ne $null) {
                $WorkInstructionURI = Get-WorkInstructionURI -Type $SelectedType
                if ($WorkInstructionURI) {
                    Edit-KanbanizeTask -TaskID $Card.taskid -BoardID $Card.BoardID -Type $SelectedType -CustomFields @{"Work Instruction"="$WorkInstructionURI"} | Out-Null
                } else {
                    Edit-KanbanizeTask -TaskID $Card.taskid -BoardID $Card.BoardID -Type $SelectedType | Out-Null
                }
            } else {
                $ToBeCreatedSelectedType = $global:ToBeCreatedTypes | Out-GridView -PassThru
                if ($ToBeCreatedSelectedType -ne $null) {
                    $global:CardsThatNeedToBeCreatedTypes += [pscustomobject]@{taskid=$Card.taskid; type=$ToBeCreatedSelectedType;BoardID=$Card.BoardID}
                } else {
                    $ToBeCreatedSelectedType = read-host "Enter the new type you want to use for this card"
                    $global:CardsThatNeedToBeCreatedTypes += [pscustomobject]@{taskid=$Card.taskid; type=$ToBeCreatedSelectedType;BoardID=$Card.BoardID}
                    $global:ToBeCreatedTypes += $ToBeCreatedSelectedType
                }
            }
        }

        if($card.color -notin ("#cc1a33","#f37325","#77569b","#067db7")) {
            $Priority = get-MultipleChoiceQuestionAnswered -Question "What priority level should this request have?" -Choices 1,2,3,4 -DefaultChoice 3
            $color = switch($Priority) {
                1 { "cc1a33" } #Red for priority 1
                2 { "f37325" } #Orange for priority 2
                3 { "77569b" } #Yello for priority 3
                4 { "067db7" } #Blue for priority 4
            }
            Write-Verbose "Color: $color"
            Edit-KanbanizeTask -BoardID $Card.BoardID -TaskID $Card.taskid -Color $color
        }

        #$WorkInstructionsForThisRequest = $card.Type -in $ApprovedWorkInstructionsInEvernote
        
        #if($WorkInstructionsForThisRequest) {
        #    $DestinationBoardID = $HelpDeskProcessBoardID
        #} else {
        #    $DestinationBoardID = $HelpDeskTechnicianProcessBoardID
        #}

        #For now send everything to the Help Desk Process board
        $DestinationBoardID = $HelpDeskProcessBoardID

        Write-Verbose "Destination column: $DestinationBoardID"

        $FurtherTriageNeeded = get-MultipleChoiceQuestionAnswered -Question "Does this need to be triaged further?" -Choices "Yes","No" | 
            ConvertTo-Boolean
        
        if ($FurtherTriageNeeded) {
                $SendEmailToRequestorForMoreInformation = get-MultipleChoiceQuestionAnswered -Question "Send email to requestor for more information?" -Choices "Yes","No" | ConvertTo-Boolean
                if ($SendEmailToRequestorForMoreInformation) {
                    Set-KanbanizeContextCard -Card $Card
                    Send-MailMessageToRequestor -DaysToWaitForResponseBeforeFollowUp 3
                }
        } else {
            $NeedToBeEscalated = get-MultipleChoiceQuestionAnswered -Question "Does this need to be escalated?" -Choices "Yes","No" | 
            ConvertTo-Boolean
        
            if($NeedToBeEscalated) {
                $DestinationLane = "Unplanned Work"
                Move-KanbanizeTask -BoardID $DestinationBoardID -TaskID $Card.taskid -Lane $DestinationLane -Column "Requested.Ready to be worked on"
            } else { 
                $DestinationLane = "Planned Work"
                Move-KanbanizeTask -BoardID $DestinationBoardID -TaskID $Card.taskid -Lane $DestinationLane -Column "Requested.Ready to be worked on"

                <#
                $CardsThatNeedToBeSorted = $Cards | 
                where {$_.columnpath -eq $DestinationColumn -and $_.lanename -eq "Planned Work"} |
                sort positionint

                $SortedCards = $CardsThatNeedToBeSorted |
                sort priorityint, trackitid
                $PositionOfTheLastCardInTheSamePriortiyLevel = $SortedCards |
                    where priorityint -EQ $(if($Card.PriorityInt){$Card.PriorityInt}else{$Priority}) |
                    select -Last 1 -ExpandProperty PositionInt
            
                $RightPosition = if($PositionOfTheLastCardInTheSamePriortiyLevel) {
                    $PositionOfTheLastCardInTheSamePriortiyLevel + 1
                } else { 0 }
                Write-Verbose "Rightposition in column: $RightPosition"
            
                Move-KanbanizeTask -BoardID $DestinationBoardID -TaskID $Card.taskid -Lane $DestinationLane -Column $DestinationColumn -Position $RightPosition
                #>
            }
        }

        Write-Verbose "DestinationLane: $DestinationLane"
    }

    $global:ToBeCreatedTypes
    Read-Host "Create types in Kanbanize for all the types listed above and then hit enter"

    $global:CardsThatNeedToBeCreatedTypes
    $global:CardsThatNeedToBeCreatedTypes | % {
        Edit-KanbanizeTask -TaskID $_.taskid -BoardID $_.BoardID -Type $_.type
    }
}

Function Get-NextCardToWorkOn {
    param (
        [Parameter(ParameterSetName="AssignedToMe")][Switch]$AssignedToMe
    )
    if ($AssignedToMe) {
        $NextCardToWorkOn = Get-CardsAssignedToMe | 
        Out-GridView -PassThru
    } else {
        $LoggedOnUsersName = Get-LoggedOnUserName

        $CardsAvailableToWorkOn = Get-KanbanizeTervisHelpDeskCards -HelpDeskProcess |
        where {
            $_.assignee -eq "None" -or 
            $_.assignee -eq $LoggedOnUsersName 
        } |
        where columnname -In "Waiting to be worked on", "Ready to be worked on"

        $CardsInUnplannedWorkWaitingOrReadyToBeWorkedOn = $CardsAvailableToWorkOn | 
        where lanename -eq "Unplanned Work"

        if ($CardsInUnplannedWorkWaitingOrReadyToBeWorkedOn) {
            $CardsInLane = $CardsInUnplannedWorkWaitingOrReadyToBeWorkedOn
        } else {
            $CardsInLane = $CardsAvailableToWorkOn | 
            where lanename -eq "Planned Work"
        }

        if ($CardsInLane | where columnname -eq "Waiting to be worked on") {
            $CardsInColumn = $CardsInLane | where columnname -eq "Waiting to be worked on"
        } else {
            $CardsInColumn = $CardsInLane | where columnname -eq "Ready to be worked on"
        }
    
        $NextCardToWorkOn = $CardsInColumn |
        sort positionint |
        select -First 1
    }
    
    Start-WorkingOnCard -Card $NextCardToWorkOn
}

Function Get-CardBeingWorkedOn {
    Get-KanbanizeTervisHelpDeskCards -HelpDeskProcess |
    where columnname -eq "Being worked on" |
    where assignee -eq $(Get-LoggedOnUserName)
}

Function Get-CardDetails {
    $Card = (Get-KanbanizeContextCard)
    if (-Not $Card) {
        $Card = (Get-CardBeingWorkedOn | Out-GridView -PassThru)
    }

    Get-TervisWorkOrderDetails -Card $Card
}

Function Get-CardsAssignedToMe {
    Get-KanbanizeTervisHelpDeskCards -HelpDeskProcess |
    where columnname -In "Waiting to be worked on","Ready to be worked on" |
    where assignee -eq $(Get-LoggedOnUserName)
}

Function Start-WorkingOnCard {
    param (
        [Parameter(ParameterSetName="Card")]$Card
    )
    Set-KanbanizeContextCard $Card
    Move-KanbanizeTask -BoardID $Card.BoardID -TaskID $Card.taskid -Column "Being worked on" | Out-Null
    Edit-KanbanizeTask -BoardID $Card.BoardID -TaskID $Card.taskid -Assignee $(Get-LoggedOnUserName) | Out-Null
    Get-TervisWorkOrderDetails -Card $Card
}

Function Stop-WorkingOnCard {
    param (
        $Card = (Get-KanbanizeContextCard),
        [Switch]$LeaveAssignedToMe
    )
    
    Move-KanbanizeTask -BoardID $Card.BoardID -TaskID $Card.taskid -Column "Waiting to be worked on" -Position 0 | Out-Null
    
    if (-not $LeaveAssignedToMe) {
        Edit-KanbanizeTask -BoardID $Card.BoardID -TaskID $Card.taskid -Assignee "None" | Out-Null
    }

    Remove-KanbanizeContextCard
}

Function Set-KanbanizeContextCard {
    Param (
        [Parameter(ValueFromPipeline)]$Card = (Get-CardBeingWorkedOn | Out-GridView -PassThru)
    )
    if ($Card) {
        $Global:KanbanizeContextCard = $Card
    }
}

Function Get-KanbanizeContextCard {
    $Global:KanbanizeContextCard
}

Function Remove-KanbanizeContextCard {
    Remove-item Variable:Global:KanbanizeContextCard
}

Function Send-MailMessageToRequestor {
    [CmdletBinding(DefaultParameterSetName="NeedToWait")]
    param (        
        [Parameter(Mandatory, ParameterSetName="NeedToWait")]
        [ValidateRange(1,14)]
        $DaysToWaitForResponseBeforeFollowUp,        
        
        [Parameter(ParameterSetName="DontNeedToWait")]
        [Switch]$DontNeedToWaitForResponse,

        [Parameter(ParameterSetName="DontNeedToWait")]
        [Parameter(ParameterSetName="NeedToWait")]        
        [Switch]$UseTemplate
    )
    $Card = Get-KanbanizeContextCard
    $WorkOrder = Get-TervisTrackITWorkOrder -WorkOrderNumber $Card.TrackITID

    $HelpDeskSignature = "`r`n`r`nThanks,`r`n`r`nIT Help Desk"

    $DefaultMessageText = if ($UseTemplate) {
        $ProcessedMailMessage = Get-MailMessageTemplateFile | Invoke-ProcessTemplateFile
        $ProcessedMailMessage + $HelpDeskSignature
    } else {
        "$($WorkOrder.RequestorFirstName),`r`n`r`n" + $HelpDeskSignature
    }

    $Body = Read-MultiLineInputBoxDialog -WindowTitle "Mail Message" -Message "Enter the message that will be sent" -DefaultText $DefaultMessageText
    if (-not $Body) { break }
        
    $Subject = "Re: $($Card.title) {$($Card.taskid)}"
    $Cc = "tervis_notifications@kanbanize.com"

    Send-TervisMailMessage -To $WorkOrder.RequestorEmailAddress -From HelpDeskTeam@tervis.com -Subject $Subject -Cc $Cc -Body $Body
    Start-Job -Name "Mail $($Card.taskid)" -ArgumentList $WorkOrder,$Card,$DaysToWaitForResponseBeforeFollowUp -ScriptBlock {
        param ($WorkOrder,$Card,$DaysToWaitForResponseBeforeFollowUp)
        sleep 60
        Edit-KanbanizeTask -BoardID $Card.BoardID -TaskID $Card.taskid -CustomFields @{"Scheduled Date"=(Get-Date).AddDays($DaysToWaitForResponseBeforeFollowUp).ToString("yyyy-MM-dd")}
        Move-KanbanizeTask -BoardID $Card.BoardID -TaskID $Card.taskid -Column "Waiting for scheduled date" | Out-Null
    }
}

Function Get-MailMessageTemplatePath {
    $DNSRoot = Get-ADDomain | select -ExpandProperty DNSRoot
    "\\$DNSRoot\applications\PowerShell\Production\TervisHelpDeskWorkOrderApplication\MailMessageTemplate"
}

Function Get-MailMessageAvailableProperites {
    $Card = Get-KanbanizeContextCard
    $Card
    $WorkOrder = Get-TervisTrackITWorkOrder -WorkOrderNumber $Card.TrackITID
    $WorkOrder
}

Function New-MailMessageTemplateFile {
    param (
        [Parameter(Mandatory)]$Name,
        $Type = (get-TervisKanbanizeTypes | Out-GridView -PassThru)
    )
    $MailMessageTemplateContent = Read-MultiLineInputBoxDialog -WindowTitle "Template" -Message "Enter the template content"
    if (-not $MailMessageTemplateContent) { break }

    $MailMessageTemplateContent | Out-File -FilePath "$(Get-MailMessageTemplatePath)\$Name.PSTemplate"

    $Type | New-MailMessageTemplateSymbolicLinks -MailMessageTemplateFile "$(Get-MailMessageTemplatePath)\$Name.PSTemplate"
}

Function Edit-MailMessageTemplateFile {
    $TemplateFile = Get-MailMessageTemplateFile
    $Content = Read-MultiLineInputBoxDialog -WindowTitle "Mail Message" -Message "Enter the message note that will be sent to the user" -DefaultText $($TemplateFile | Get-Content -Raw)
    if (-not $Content) { break }
    Set-MailMessageTemplateFile -File $TemplateFile -Content $Content
}

Function Get-MailMessageTemplateFile {
    param (
        $Type = $(
            $ContextCardType = Get-KanbanizeContextCard | select -ExpandProperty Type
            if ($ContextCardType) {
                $ContextCardType
            } else {
                Get-ChildItem -Path  "$(Get-MailMessageTemplatePath)" -Directory | 
                select -ExpandProperty name | 
                Out-GridView -PassThru
            }
        )
    )
    Get-ChildItem -Path "$(Get-MailMessageTemplatePath)\$Type" | Out-GridView -PassThru
}

Function Set-MailMessageTemplateFile {
    param (
        $File,
        $Content
    )
    $FinalFile = if ($File.LinkType -eq "SymbolicLink") {
        $File.Target -creplace "UNC\\","\\" | Get-Item
    } else {
        $File
    }

    $Content | Out-File -FilePath $FinalFile
}

Function New-MailMessageTemplateSymbolicLinks {
    param (
        [Parameter(Mandatory)][System.IO.FileInfo]$MailMessageTemplateFile,
        [Parameter(ValueFromPipeline)]$Type
    )
    process {
        New-Item -Path "$(Get-MailMessageTemplatePath)\$Type\$($MailMessageTemplateFile.Name)" -ItemType SymbolicLink -Value $MailMessageTemplateFile -Force
    }
}

Function New-MailMessageTemplateFileToTypeAssociation {
    Get-MailMessageTemplateFile | Add-MailMessageTemplateFileToType
}

Function Add-MailMessageTemplateFileToType {
    param (
        [Parameter(Mandatory, ValueFromPipeline)][System.IO.FileInfo]$MailMessageTemplateFile,
        $Type = (Get-KanbanizeContextCard | select -ExpandProperty Type)
    )
    process {
        New-MailMessageTemplateSymbolicLinks -MailMessageTemplateFile $MailMessageTemplateFile -Type $Type
    }
}

Function Send-MailMessageToRequestorViaOutlook {
    param (
        [ValidateRange(1,14)]$DaysToWaitForResponseBeforeFollowUp,
        [Switch]$CanPerformNextActionsWithoutResponse
    )
    $Card = Get-KanbanizeContextCard
    $WorkOrder = Get-TervisTrackITWorkOrder -WorkOrderNumber $Card.TrackITID

    Start $(New-MailToURI -To $WorkOrder.RequestorEmailAddress -Subject "Re: $($Card.title) {$($Card.taskid)}" -Cc tervis_notifications@kanbanize.com)
    
    if ($DaysToWaitForResponseBeforeFollowUp) {
        Read-Host "Press enter when message has been sent"
        Start-Job -Name "Mail $($Card.taskid)" -ArgumentList $WorkOrder,$Card,$DaysToWaitForResponseBeforeFollowUp -ScriptBlock {
            param ($WorkOrder,$Card,$DaysToWaitForResponseBeforeFollowUp)
            sleep 60
            Edit-KanbanizeTask -BoardID $Card.BoardID -TaskID $Card.taskid -CustomFields @{"Scheduled Date"=(Get-Date).AddDays($DaysToWaitForResponseBeforeFollowUp).ToString("yyyy-MM-dd")}
            Move-KanbanizeTask -BoardID $Card.BoardID -TaskID $Card.taskid -Column "Waiting for scheduled date" | Out-Null
        }
    }
}

Function Close-WorkOrder {
    $Card = Get-KanbanizeContextCard

    if ($Card.Type -ne "None") {
        $TypeCorrect = get-MultipleChoiceQuestionAnswered -Question "Type ($($Card.Type)) correct?" -Choices "Yes","No" | ConvertTo-Boolean               
    }
        
    if (-not $TypeCorrect -or ($Card.Type -eq "None")) {        
        $SelectedType = get-TervisKanbanizeTypes | Out-GridView -PassThru
    
        if ($SelectedType -ne $null) {
            Edit-KanbanizeTask -TaskID $Card.taskid -BoardID $Card.BoardID -Type $SelectedType | Out-Null
        } else {
            Throw "Please create the missing type that should be assigned to this work order and then close the work order again"
        }
    }

    if ($Card.TrackITID) {
        $WorkOrder = Get-TervisTrackITWorkOrder -WorkOrderNumber $Card.TrackITID

        $DefaultCloseMessage = "$($WorkOrder.RequestorFirstName),`r`n`r`n`r`n`r`nIf you have any further issues please give us a call at 2248 or 941-441-3168`r`n`r`nThanks,`r`n`r`nIT Help Desk"
        $Resolution = Read-MultiLineInputBoxDialog -WindowTitle "Resolution" -Message "Enter the final resolution note that will be sent to the user" -DefaultText $DefaultCloseMessage
        if (-not $Resolution) { break }
        
        Import-module TrackItWebAPIPowerShell -Force #Something is broken as this line shouldn't be required but it is
        Invoke-TrackITLogin -Username helpdeskbot -Pwd helpdeskbot
        $Response = Close-TrackITWorkOrder -WorkOrderNumber $Card.TrackITID -Resolution $Resolution
        if (-not ($Response.success | ConvertTo-Boolean)) { 
            Throw "Closing the track it work order failed. $($Response.data)" 
        }
        
        $Subject = "Re: $($Card.title) {$($Card.taskid)}"
        $Cc = "tervis_notifications@kanbanize.com"
        Send-TervisMailMessage -To $WorkOrder.RequestorEmailAddress -From HelpDeskTeam@tervis.com -Subject $Subject -Cc $Cc -Body $Resolution
    } else {
        $Requestor = if ($Card.Requestor) {$Card.Requestor} else {$Card.Reporter}
        $DefaultCloseMessage = "@$($Requestor),`r`n`r`n`r`n`r`nIf you have any further issues please give us a call at 2248 or 941-441-3168`r`n`r`nThanks,`r`n`r`nIT Help Desk"
        $Resolution = Read-MultiLineInputBoxDialog -WindowTitle "Resolution" -Message "Enter the final resolution note that will be sent to the user" -DefaultText $DefaultCloseMessage
        if (-not $Resolution) { break }

        Add-KanbanizeComment -TaskID $Card.TaskID -Comment $Resolution
    }

    Move-KanbanizeTask -BoardID $Card.BoardID -TaskID $Card.taskid -Column "Done" | Out-Null
    Remove-KanbanizeContextCard
}

Function Invoke-OpenWorkOrderInTrackIT {
    $Card = Get-KanbanizeContextCard
    Start $($Card.customfields | where name -eq trackiturl | select -ExpandProperty value)
}