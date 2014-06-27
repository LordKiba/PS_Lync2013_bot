<#	
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2014 v4.1.54
	 Created on:   	6/25/2014 7:20 PM
	 Created by:   	James A Kulikowski
	 Organization: 	GFG
	 Filename:     	Lync Bot.psm1
	-------------------------------------------------------------------------
	 Module Name: Lync Bot
	===========================================================================
#>

# uncomment next line for error supression
#$ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue

# Clear all previous subscribed events
Get-EventSubscriber | Unregister-Event

$ModelPaths = @(); $ModelPaths = "C:\Program Files\Microsoft Office\Office15\LyncSDK\Assemblies\Desktop\Microsoft.Lync.Model.DLL,C:\Program Files\Microsoft Office\Office15\LyncSDK\Assemblies\Desktop\Microsoft.Lync.Controls.DLL"; $ModelPaths = $ModelPaths.Split(",")
Foreach ($ModelPath in $ModelPaths)
{
	If (Test-Path $ModelPath)
	{
		Import-Module $ModelPath
	}
	Else
	{
		$LIBName = $ModelPath.Split("\")
		
		Write-Host "Please Import the " $LIBName[-1] "from the MS Lync SDK Before using this Module"
		break
	}
}

# Obtain entry points to the Lync.Model API for Client and Automation + error handeling
try
{
	$global:Client = [Microsoft.Lync.Model.LyncClient]::GetClient()
	
	if ($Client -eq $null)
	{
		throw "Unable to obtain client interface"
	}
	
}
catch [Microsoft.Lync.Model.ClientNotFoundException]
{
	throw "Lync client is not running! Please launch your Lync client."
}
$Client.add_StateChanged
$global:Auto = [Microsoft.Lync.Model.LyncClient]::GetAutomation()
$global:Self = $client.Self

#Test Client State for Logon/init state
function lync-state-change
{
	$Lyncstate = $Client.State
	if ($Lyncstate -eq [Microsoft.Lync.Model.ClientState]::Uninitialized)
	{
		Write-Host "Lync Client not in Initialized State.`n Initializeing ..."
		$ar = $Client.BeginInitialize()
		$Client.EndInitialize($ar)
		
	}
	elseif ($Lyncstate -eq [Microsoft.Lync.Model.ClientState]::SignedIn)
	{
		$usr = $Self.Contact.Uri.ToString().TrimStart("sip:")
		Write-Host "User is logged in and Powershell is ready for Bot Startup"
		function Global:prompt { "Lync Bot CLI [$usr] >" }
		prompt
	}
	elseif ($Lyncstate -eq [Microsoft.Lync.Model.ClientState]::SignedOut)
	{
		Write-Host "No user is logged into the Lync Client.`n login before running Lync-Bot"
	}
	elseif ($Lyncstate -eq [Microsoft.Lync.Model.ClientState]::SigningIn)
	{
		Write-Host "Client is Logging in.`n Please standby."
	}
	elseif ($Lyncstate -eq [Microsoft.Lync.Model.ClientState]::ShuttingDown)
	{
		Write-Host "Lync Client is Shutting Down.`n Terminateing Bot Event Handelers"
	}
}

function lync-send-msg($msg)
{
	# Send the message
	$null = $Modality.BeginSendMessage($msg, $null, $msg)
}

function Lync-Availability
{
	Param (
		[ValidateSet("Appear Offline", "Available", "Away", "Busy", "Do Not Disturb", "Be Right Back", "Off Work")]
		[string]
		$Availability,
		# ActivityId as string
		[string]
		$ActivityId,
		# String value to be configured as personal note in the Lync client
		[string]
		$PersonalNote,
		# String value to be configured as location in the Lync client
		[string]
		$Location
	)
	$ContactInfo = New-Object 'System.Collections.Generic.Dictionary[Microsoft.Lync.Model.PublishableContactInformationType, object]'
	
	switch ($Availability)
	{
		"Available" { $AvailabilityId = 3000 }
		"Appear Offline" { $AvailabilityId = 18000 }
		"Away" { $AvailabilityId = 15000 }
		"Busy" { $AvailabilityId = 6000 }
		"Do Not Disturb" { $AvailabilityId = 9000 }
		"Be Right Back" { $AvailabilityId = 12000 }
		"Off Work" { $AvailabilityId = 15500 }
	}
	
	if ($Availability)
	{
		$ContactInfo.Add([Microsoft.Lync.Model.PublishableContactInformationType]::Availability, $AvailabilityId)
	}
	
	if ($ActivityId)
	{
		$ContactInfo.Add([Microsoft.Lync.Model.PublishableContactInformationType]::ActivityId, $ActivityId)
	}
	
	if ($PersonalNote)
	{
		$ContactInfo.Add([Microsoft.Lync.Model.PublishableContactInformationType]::PersonalNote, $PersonalNote)
	}
	
	if ($Location)
	{
		$ContactInfo.Add([Microsoft.Lync.Model.PublishableContactInformationType]::LocationName, $Location)
	}
	
	if ($ContactInfo.Count -gt 0)
	{
		
		$Publish = $Self.BeginPublishContactInformation($ContactInfo, $null, $null)
		$self.EndPublishContactInformation($Publish)
		
	}
	else
	{
		
		Write-Warning "No options supplied, no action was performed"
		
	}
	
	
}

# Job that is called on new message recievedvevent handeler
$global:action = {
	
	# get the conversation that caused the event
	$Conversation = $Event.Sender.Conversation
	
	Write-Host $Event.SourceArgs | fl -f 
	
	# Create a new msg collection for the response
	$msg = New-Object "System.Collections.Generic.Dictionary[Microsoft.Lync.Model.Conversation.InstantMessageContentType,String]"
	# Modality Type
	$Global:Modality = $Conversation.Modalities[1]
	
	# The message recieved
	[string]$msgStr = $Event.SourceArgs.Text
	$msgStr = $msgStr.ToString().ToLower().Trim()
	Write-Host $msgStr
	$BotCMD = $msgStr.Split(" ")
	$BotCMD = $BotCMD[0].TrimEnd(".!?")
	Write-Host "First index in array is : " $BotCMD
	
	# switch commands / messages - add what you like here
	switch ($BotCMD)
	{
		"hey" {
			$sendMe = 1
			$msg.Add(0, 'Hello')
			
		}
		"hi" {
			$sendMe = 1
			$msg.Add(0, 'Hello')
			
		}
		"hello" {
			$sendMe = 1
			$msg.Add(0, 'Hey Whats up?')
			
		}
		"moo" {
			$sendMe = 1
			$msg.Add(0, 'Are you a cow?')	
		}
		"help" {
			$sendMe = 1
			$msg.Add(0, 'Mr James bot at your service!')
			lync-send-msg -msg $msg
			sleep -Milliseconds 250
			$msg = New-Object "System.Collections.Generic.Dictionary[Microsoft.Lync.Model.Conversation.InstantMessageContentType,String]"
			$msg.Add(0, 'available commands:')
			lync-send-msg -msg $msg
			sleep -Milliseconds 250
			$msg = New-Object "System.Collections.Generic.Dictionary[Microsoft.Lync.Model.Conversation.InstantMessageContentType,String]"
			$msg.Add(0, 'hi, hey, hello')
			lync-send-msg -msg $msg
			sleep -Milliseconds 250
			$msg = New-Object "System.Collections.Generic.Dictionary[Microsoft.Lync.Model.Conversation.InstantMessageContentType,String]"
			$msg.Add(0, 'moo')
			lync-send-msg -msg $msg
			sleep -Milliseconds 250
			$msg = New-Object "System.Collections.Generic.Dictionary[Microsoft.Lync.Model.Conversation.InstantMessageContentType,String]"
			$msg.Add(0, 'time')
			lync-send-msg -msg $msg
			sleep -Milliseconds 250
			$msg = New-Object "System.Collections.Generic.Dictionary[Microsoft.Lync.Model.Conversation.InstantMessageContentType,String]"
			$msg.Add(0, '!busy, !free, !brb, !offline, !away')
		}
		"time" {
			$sendMe = 1
			$now = Get-Date
			$msg.Add(0, 'Current Date and Time : ' + $now)
		}
		"!busy" {
			$sendMe = 1
			$date = [DateTime]::Now
			Lync-Availability -Availability 'Busy' -Location 'CyberSpace' -PersonalNote "Set availability to Busy using the Lync Model API in PowerShell on $date"
			$msg.Add(0, 'Bot Command Accepted: setting Avaiability to Busy')
		}
		"!free" {
			$sendMe = 1
			$date = [DateTime]::Now
			Lync-Availability -Availability 'Available'  -Location 'CyberSpace' -PersonalNote "Set availability to Available using the Lync Model API in PowerShell on $date"
			$msg.Add(0, 'Bot Command Accepted: setting Avaiability to Available')
		}
		"!brb" {
			$sendMe = 1
			$date = [DateTime]::Now
			Lync-Availability -Availability 'Be Right Back' -Location 'Away From Keybord' -PersonalNote "Be Right Back"
			$msg.Add(0, 'Bot Command Accepted: setting Avaiability to Be Right Back')
		}
		"!away" {
			$sendMe = 1
			$date = [DateTime]::Now
			Lync-Availability -Availability 'Away' -Location 'Away From Keybord' -PersonalNote "I'm not here at the moment."
			$msg.Add(0, 'Bot Command Accepted: setting Avaiability to away')
		}
		"!offline" {
			$sendMe = 1
			$date = [DateTime]::Now
			Lync-Availability -Availability 'Appear Offline' 
			$msg.Add(0, 'Bot Command Accepted: setting Avaiability to Offline')
		}
		"dcpromo"{
			$sendMe = 1
			$msg.Add(0, 'Promoteing Server as new node in Forest Gentry.IsMyHero.Local')
		}
		"thanks"{ $sendMe = 1
			$msg.Add(0, 'NP, Your Welcome')
		}
		
		"thank"{
			$sendMe = 1
			$msg.Add(0, 'NP, Your Welcome')
		}
		default
		{
			# do nothing
			$sendMe = 0
		}
	}
	
	if ($sendMe -eq 1)
	{
		# Send the message
		lync-send-msg -msg $msg
	}
}

function Lync-NoBot
{
	$usr = $Self.Contact.Uri.ToString().TrimStart("sip:")
	# Clear all Bot subscribed events
	Get-EventSubscriber | Unregister-Event
	function Global:prompt { "Lync Bot CLI [$usr] >" }
}

function Lync-Signout
{
	Write-Host "Unsubscribeing All Lync Events this session"
	Lync-NoBot
	Write-Host "initializeing signout process"
	$ar = $Client.BeginSignOut($communicatorClientCallback, $null)
	while ($ar.IsCompleted -eq $false) { }
	$Client.EndSignOut($ar)
	Write-Host "Signed out of Lync Client"
	function Global:prompt { "Lync Bot CLI >" }
prompt	
}

function Lync-Shutdown ([system.Boolean]$Confirm)
{
	if ($Confirm -eq $true)
	{
		Lync-Signout
		Write-Host "Starting Client Shutdown"
		$ar = $Client.BeginShutdown($communicatorClientCallback, $null)
		while ($ar.IsCompleted -eq $false) { }
		$Client.EndShutdown($ar)
		Write-Host "Client Shutdown: Background Process Terminated."
	}
	else { Write-Host 'Please confirm Client Shutdown with "-Confirm"' }
	
	
}

function Lync-Bot
{
	# Register events for current open conversation participants
	foreach ($con in $client.ConversationManager.Conversations)
	{
		# For each participant in the conversation
		$moo = $con.Participants | Where { !$_.IsSelf }
		
		foreach ($mo in $moo)
		{
			$mo.Contact.uri
			if (!(Get-EventSubscriber $mo.Contact.uri))
			{
				Register-ObjectEvent -InputObject $mo.Modalities[1] `
									 -EventName "InstantMessageReceived" `
									 -SourceIdentifier $mo.Contact.uri `
									 -action $action
		}
	}
}

# Add event to pickup new conversations and register events for new participants
	$conversationMgr = $client.ConversationManager
	Register-ObjectEvent -InputObject $conversationMgr `
						 -EventName "ConversationAdded" `
						 -SourceIdentifier "NewIncomingConversation" `
						 -action {
		$client = [Microsoft.Lync.Model.LyncClient]::GetClient()
		foreach ($con in $client.ConversationManager.Conversations)
		{
			# For each participant in the conversation
			$moo = $con.Participants | Where { !$_.IsSelf }
			foreach ($mo in $moo)
			{
				$mo.Contact.uri
				if (!(Get-EventSubscriber $mo.Contact.uri))
				{
					Register-ObjectEvent -InputObject $mo.Modalities[1] `
										 -EventName "InstantMessageReceived" `
										 -SourceIdentifier $mo.Contact.uri `
										 -action $action
				}
			}
		}
	}
	$usr = $Self.Contact.Uri.ToString().TrimStart("sip:")
	function Global:prompt { "Lync Bot CLI [$usr] [Bot:ON] >" }
}

# Runtime INIT
clear
$Host.UI.RawUI.WindowTitle = "Lync Bot C&C Console"
function Global:prompt { "Lync Bot CLI >" }
prompt

#State change notification event registration
Register-ObjectEvent -InputObject $Client `
					 -EventName "StateChanged" `
					 -SourceIdentifier "LyncClientStateChanged" `
					 -action { lync-state-change }

#init state change processing
lync-state-change

function Lync-reload
{
	Lync-NoBot
	Remove-Module 'Lync Bot'
	Import-Module 'Lync Bot'
}

# export module members
Export-ModuleMember lync-State-Change
Export-ModuleMember lync-send-msg
Export-ModuleMember Lync-Bot
Export-ModuleMember Lync-NoBot
Export-ModuleMember Lync-Shutdown
Export-ModuleMember Lync-Signout
Export-ModuleMember Lync-Availability
Export-ModuleMember Lync-reload