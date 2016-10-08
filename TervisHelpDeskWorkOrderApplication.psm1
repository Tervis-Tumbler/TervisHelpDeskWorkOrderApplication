﻿#Requires -Modules KanbanizePowerShell, TrackITUnOfficial, TrackITWebAPIPowerShell, get-MultipleChoiceQuestionAnswered, TervisTrackITWebAPIPowerShell
#Requires -Version 4

function Invoke-PrioritizeConfirmTypeAndMoveCard {
    [CmdletBinding()]
    param()

    #$VerbosePreference = "continue"

    Import-Module KanbanizePowerShell -Force
    Import-module TrackItWebAPIPowerShell -Force

    Invoke-TrackITLogin -Username helpdeskbot -Pwd helpdeskbot

    $KanbanizeBoards = Get-KanbanizeProjectsAndBoards

    $HelpDeskProcessBoardID = $KanbanizeBoards.projects.boards | where name -EQ "Help Desk Process" | select -ExpandProperty ID
    $HelpDeskTechnicianProcessBoardID = $KanbanizeBoards.projects.boards | where name -EQ "Help Desk Technician Process" | select -ExpandProperty ID

    $Types = get-TervisKanbanizeTypes

    #$WaitingToBePrioritized = Get-KanbanizeTervisHelpDeskCards -HelpDeskProcess |
    #where columnpath -Match "Waiting to be prioritized" |
    #sort positionint

    $WaitingToBePrioritized = Get-KanbanizeTervisHelpDeskCards -HelpDeskTriageProcess |
    where columnpath -NotMatch "Waiting for scheduled date" |
    sort positionint

    $global:CardsThatNeedToBeCreatedTypes = @()
    $global:ToBeCreatedTypes = @()

    foreach ($Card in $WaitingToBePrioritized) {
        Get-TervisWorkOrderDetails -Card $Card

        read-host "Hit enter once you have reviewed the details about this request"

        if ($Card.Type -ne "None") {
            $TypeCorrect = get-MultipleChoiceQuestionAnswered -Question "Type correct?" -Choices "Yes","No" | ConvertTo-Boolean               
        }
        
        if (-not $TypeCorrect -or ($Card.Type -eq "None")) {        
            $SelectedType = $Types | Out-GridView -PassThru
    
            if ($SelectedType -ne $null) {
                Edit-KanbanizeTask -TaskID $Card.taskid -BoardID $Card.BoardID -Type $SelectedType        
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

        #For now send everything tot he Help Desk Process board
        $DestinationBoardID = $HelpDeskProcessBoardID

        Write-Verbose "Destination column: $DestinationBoardID"

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
    $LoggedOnUsersName = Get-LoggedOnUserName

    $Cards = Get-KanbanizeTervisHelpDeskCards -HelpDeskProcess |
    where {
        $_.assignee -eq "None" -or 
        $_.assignee -eq $LoggedOnUsersName 
    }

    if ($Cards | where lanename -eq "Unplanned Work") {
        $CardsInLane = $Cards | where lanename -eq "Unplanned Work"
    } else {
        $CardsInLane = $Cards | where lanename -eq "Planned Work"
    }

    if ($CardsInLane | where columnname -eq "Waiting to be worked on") {
        $CardsInColumn = $CardsInLane | where columnname -eq "Waiting to be worked on"
    } else {
        $CardsInColumn = $CardsInLane | where columnname -eq "Ready to be worked on"
    }
    
    $NextCardToWorkOn = $CardsInColumn |
    sort positionint |
    select -First 1
    
    Start-WorkingOnCard -Card $NextCardToWorkOn
}

Function Get-CardBeingWorkedOn {
    Get-KanbanizeTervisHelpDeskCards -HelpDeskProcess |
    where assignee -eq $(Get-LoggedOnUserName)
}

Function Get-CardDetails {
    $Card = (Get-KanbanizeContextCard)
    if (-Not $Card) {
        $Card = (Get-CardBeingWorkedOn | Out-GridView -PassThru)
    }

    Get-TervisWorkOrderDetails -Card $Card
}

Function Start-WorkingOnCard {
    param (
        $Card
    )
    Set-KanbanizeContextCard $Card
    Move-KanbanizeTask -BoardID $Card.BoardID -TaskID $Card.taskid -Column "Being worked on" | Out-Null
    Edit-KanbanizeTask -BoardID $Card.BoardID -TaskID $Card.taskid -Assignee $LoggedOnUsersName | Out-Null

    Invoke-TrackITLogin -Username helpdeskbot -Pwd helpdeskbot | Out-Null
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
        $Card = (Get-CardBeingWorkedOn | Out-GridView -PassThru)
    )
    $Global:KanbanizeContextCard = $Card
}

Function Get-KanbanizeContextCard {
    $Global:KanbanizeContextCard
}

Function Remove-KanbanizeContextCard {
   Remove-item Variable:Global:KanbanizeContextCard
}

Function Get-LoggedOnUserName {
    Get-aduser $Env:USERNAME | select -ExpandProperty Name
}