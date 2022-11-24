

$buffer=@{"before"=0; "after"=5}  # Buffer either side of /now/ to look for events. Note that this works on intersection of the search period and any time in the meeting.


<# Note time zones
    Get-MgUserCalendar does searches based on UTC and returns data with UTC
    So try to work with UTC for all searches and comparisons. But produce some conversions to local time for display
#>


# Get calendar details


    # Get username
    #   Considered using $env:username or (whoami /upn). But this method gets the Teams account specifically
    Get-ItemPropertyValue HKCU:\SOFTWARE\Microsoft\Office\Teams -name HomeUserUPN

    echo "Checking calendar for user: $myuserid"

    # Connect-MgGraph will need some consent
    # TODO: Consider app consent instead ?
    Connect-MgGraph -Scopes "Calendars.Read" 

    $calendarID=(Get-MgUserCalendar -UserId $myuserid).ID


    # Use calendarView. See https://docs.microsoft.com/en-us/graph/api/user-list-calendarview
    # Specify a date/time range and it returns any instances of events that intersect with that time
    # Note instances. This avoids issues with recurring events and -mgusercalendarevent
    # And intersection of times. This means we don't need to worry about length of events

    $now=(get-date).ToUniversalTime() 

    $starttime=$now.AddMinutes(-$buffer["before"])
    $endtime=$now.AddMinutes($buffer["after"])

    $events=@(Get-MgUserCalendarView -UserId $myuserid -CalendarId $calendarID -StartDateTime $starttime -EndDateTime $endtime | select *,@{N="Start_l";E={(get-date $_.Start.DateTime).ToLocalTime()}},@{N="End_l";E={(get-date $_.End.DateTime).ToLocalTime()}} | ?{ -not $_.isallday })

if ($events.count -eq 0) {
    "No events found between $starttime and $endtime"

    # Friendly reminder of whenever the next meeting is
    #    Need EndDateTime enough to find a meeting, but if we return a full page we might be missing some. So check and loop.
    $d=7 ; $p=20; do { $nextmeetings=Get-MgUserCalendarView -UserId $myuserid -CalendarId $calendarID -StartDateTime $starttime -EndDateTime $endtime.adddays($d) -PageSize $p | ?{ -not $_.isallday }; <#"d: $d   p: $p   count: $($nextmeetings.Count)";#> $d-=1 } while ($nextmeetings.count -ge $p)
    $nextmeeting=$nextmeetings | sort { $_.Start.DateTime } | select -first 1
    
    #    Output
    "`n`nNext meeting in "+$(((get-date $nextmeeting.start.DateTime)-(get-date).ToUniversalTime()) | select @{N="S";E={""+[math]::floor($_.TotalHours)+" hrs "+$_.minutes+" mins"}} | select -ExpandProperty S)
    $nextmeeting | select Subject,@{N="Start";E={(get-date $_.Start.DateTime).ToLocalTime().ToString('ddd dd/MM HH:mm')}},@{N="Who";E={($_.Attendees.EmailAddress.Name | %{ ($_ -split ' ')[0] }) -join ", "}} | fl *

} elseif ($events.count -gt 1) {
    "Multiple events found"
    $events| select Subject,Start_l,End_l | ft -AutoSize

    # Pick the nearest meeting 
    $nearestmeeting=($events | select @{N="Diff";E={[math]::Abs(($now.ToUniversalTime()-(get-date $_.Start.DateTime)).totalseconds)}},* | sort Diff)[0]

    $meetingtoopen=$nearestmeeting

} else {
    "Found 1 event: $($events[0].Subject)"

    $meetingtoopen=$events[0]

}

if ($meetingtoopen -ne $null) {
    "Opening meeting $($meetingtoopen.Subject)"


    if ( -not $meetingtoopen.IsOnlineMeeting) {
        "  Event is not an online meeting"

        if ($meetingtoopen.Organizer.EmailAddress.Address -match 'calendar.google.com$') {
            "  But appears to be a Google Meet"
            $gmeeturl=[regex]::Match($meetingtoopen.Body.Content,'<a href="(https://meet.google.com/bgz-zses-pwt\?hs=224)">').groups[1].value
            Start-Process $gmeeturl
        }

        elseif ($meetingtoopen.Location.DisplayName -match '^http(s)?://') {
            "  But appears to have a URL in Location"
            Start-Process $meetingtoopen.Location.DisplayName
        }

    } else {
        "  Event is an online meeting"
        if ( -not ($meetingtoopen.OnlineMeetingProvider -eq "teamsForBusiness") ) {
            "    Event is not a Teams meeting. Opening URL"
            $meetingurl=@(($meetingtoopen.OnlineMeeting.JoinUrl, $meetingtoopen.OnlineMeetingUrl) | sort { $_.length })[-1]
            Start-Process $meetingurl
        } else {
            "    Event is a Teams meeting"
            $meetingurl=$meetingtoopen.OnlineMeeting.JoinUrl
            $teamsmeetingurl=$meetingurl -replace 'https://teams.microsoft.com','msteams:' # Rewrite to a msteams: URI which will launch Teams directly
            
            "    Launching meeting"
            Start-Process $teamsmeetingurl

            # Now click Join Now. Can't script this, so simulate button presses
            $wsh=new-object -ComObject wscript.shell
            Start-Sleep -Seconds 1
            $wsh.AppActivate($meetingtoopen.Subject) # Find the window based on the meeting subject
            Start-Sleep -Seconds 1
            $wsh.SendKeys("+{TAB}")                  # Shift-Tab to go BACK to the "Join now" button (more consistent than going forwards)
            Start-Sleep -Seconds 1
            $wsh.SendKeys("{ENTER}")                 # Enter to select "Join now" 
        }
    }
}

Start-Sleep 15 # Wait to give chance to check output
