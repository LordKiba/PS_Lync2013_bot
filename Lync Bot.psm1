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

#region Pre-stage INIT Module Loading... 
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

# test loading of Client Automation API's 
try
{
	$global:Auto = [Microsoft.Lync.Model.LyncClient]::GetAutomation()
	
	if ($Auto -eq $null)
	{
		throw "Unable to obtain Lync Automation interface"
	}
	
}
catch
{
	throw "Automation Session is unavaiable" 
}

$global:Self = $client.Self

#endregion

#region Functions Required by Events 

function lync-send-msg($msg)
{
	Write-Host "Bot Reply : " $msg.values
	# Send the message
	$null = $Modality.BeginSendMessage($msg, $null, $msg)
}

function Lync-Availability
{
	
	<#
	.Synopsis
   		Lync-Availability is a PowerShell function to configure a set of settings in the Microsoft Lync client via the Model API.

	.DESCRIPTION
  		 The purpose of Lync-Availability is to demonstrate how PowerShell can be used to interact with the Lync SDK.

	.EXAMPLE
   		Lync-Availability -Availability Available

	.EXAMPLE
  	  Lync-Availability -Availability Away

	.EXAMPLE
    	Lync-Availability -Availability "Off Work" -ActivityId off-work

	.EXAMPLE
  	  Lync-Availability -PersonalNote test

	.EXAMPLE
  	  Lync-Availability -Availability Available -PersonalNote ("Quote of the day: " + (Get-QOTD))

	.EXAMPLE
    	Lync-Availability -Location Work

	.FUNCTIONALITY
  		 Provides a function to configure Availability, ActivityId and PersonalNote for the Microsoft Lync client.
#>
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


function do-Rot13
{
	[CmdletBinding()]
	param (
		[Parameter(
				   Mandatory = $true,
				   ValueFromPipeline = $true
		)]
		[String]
		$rot13string
	)
	
	[String] $string = $null;
	$rot13string.ToCharArray() |
	ForEach-Object {
		Write-Verbose "$($_): $([int] $_)"
		if ((([int] $_ -ge 97) -and ([int] $_ -le 109)) -or (([int] $_ -ge 65) -and ([int] $_ -le 77)))
		{
			$string += [char] ([int] $_ + 13);
		}
		elseif ((([int] $_ -ge 110) -and ([int] $_ -le 122)) -or (([int] $_ -ge 78) -and ([int] $_ -le 90)))
		{
			$string += [char] ([int] $_ - 13);
		}
		else
		{
			$string += $_
		}
	}
	$string
}

#endregion

#region Event actions
# Job that is called on new message recievedvevent handeler
$global:action = {
	
	# get the conversation that caused the event
	$Conversation = $Event.Sender.Conversation
	 
	# Create a new msg collection for the response
	$msg = New-Object "System.Collections.Generic.Dictionary[Microsoft.Lync.Model.Conversation.InstantMessageContentType,String]"
	# Modality Type
	$Global:Modality = $Conversation.Modalities[1]
	
	# The message recieved
	[string]$msgStr = $Event.SourceArgs.Text
	$msgStr = $msgStr.ToString().ToLower().Trim()
	Write-Host "message Recieved" $msgStr
	$BotCMD = $msgStr.Split(" ")
	$attribs = $msgStr.TrimStart("$BotCMD[0]")
	$BotCMD = $BotCMD[0].TrimEnd(".!?")
#	Write-Host "First index in array is : " $BotCMD
	
	# switch commands / messages - add what you like here
	switch ($BotCMD)
	{
		"sorry" {
			$sendMe = 1
			$msg.Add(0, 'you should be sorry for what you have done :-(')
			
		}
		"wassup" {
			$sendMe = 1
			$msg.Add(0, 'nothing much, wassup with you?')
			
		}
		"how" {
			$sendMe = 1
			$msg.Add(0, 'How about you google it :-)')
			
		}
		"yes" {
			$sendMe = 1
			$msg.Add(0, 'Yessssssssiiirrrrrrrr')
			
		}
		"I" {
			$sendMe = 1
			$msg.Add(0, 'There is no I in team!')
			
		}
		"sup" {
			$sendMe = 1
			$msg.Add(0, 'Is it time for supper already?')
			
		}
		"you" {
			$sendMe = 1
			$msg.Add(0, 'its always you you you, what about me?')
			
		}
		"lol" {
			$sendMe = 1
			$msg.Add(0, 'whats so funny?')
			
		}
		"just" {
			$sendMe = 1
			$msg.Add(0, 'just what?')
			
		}
		"yeah" {
			$sendMe = 1
			$msg.Add(0, 'yeah what?')
			
		}
		"yea" {
			$sendMe = 1
			$msg.Add(0, 'yea what?')
			
		}
		"nothing" {
			$sendMe = 1
			$msg.Add(0, 'sounds like something!')
			
		}
		"what" {
			$sendMe = 1
			$msg.Add(0, 'whatup with you?')
			
		}
		"what`'s" {
			$sendMe = 1
			$msg.Add(0, "what`'s up with you?")
			
		}
		"whats" {
			$sendMe = 1
			$msg.Add(0, 'whats up with you?')
			
		}
		"nm" {
			$sendMe = 1
			$msg.Add(0, 'just chillin...')
			
		}
		"yo" {
			$sendMe = 1
			$msg.Add(0, 'Wassup?')
			
		}
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
		"huh"{
			$sendMe = 1
			$msg.Add(0, 'What ?')
		}
		"md5"{
			$md5 = new-object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
			$utf8 = new-object -TypeName System.Text.UTF8Encoding
			$hash = [System.BitConverter]::ToString($md5.ComputeHash($utf8.GetBytes($attribs)))
			$hash = $hash -replace "-", ""
			$sendMe = 1
			$msg.Add(0, "$hash")
		}
		"rot13"{
			$out = do-Rot13 $attribs
			$sendMe = 1
			$msg.Add(0, "$out")
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

#endregion

#region Lync Bot Management Functions
#Test Client State for Logon/init state
function lync-state-change
{
		<#
		.SYNOPSIS
			lync-state-change is a PowerShell function to detect the current Lync Client state.  
		
		.DESCRIPTION
   			The purpose of lync-state-change is to demonstrate how PowerShell can be used to interact with the Lync SDK.
	
		.EXAMPLE
			Lync-reload
	#>
	
	
	$Lyncstate = $Client.State
	if ($Lyncstate -eq [Microsoft.Lync.Model.ClientState]::Uninitialized)
	{
		Write-Host "Lync Client not in Initialized State.`n Initializeing ..."
		$ar = $Client.BeginInitialize()
		$Client.EndInitialize($ar)
		
	}
	elseif ($Lyncstate -eq [Microsoft.Lync.Model.ClientState]::SignedIn)
	{
		Write-Host "User is logged in and Powershell is ready for Bot Startup"
		function Global:prompt { [System.String]$Global:usr = $Client.Uri.TrimStart("sip:"); "Lync Bot CLI [$usr] >" }
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

function Lync-NoBot
{
	<#
		.SYNOPSIS
			Lync-NoBot is a PowerShell function to turn off the Lync Autoresponce bot 

		.DESCRIPTION
   			The purpose of Lync-NoBot is to demonstrate how PowerShell can be used to interact with the Lync SDK.
	
		.EXAMPLE
			Lync-NoBot
	#>
	
	# Clear all Bot subscribed events
	Get-EventSubscriber | Unregister-Event
	function Global:prompt { [System.String]$Global:usr = $Client.Uri.TrimStart("sip:");"Lync Bot CLI [$usr] >" }
}

function Lync-Signout
{
		<#
		.SYNOPSIS
			Lync-Signout is a PowerShell function to turn off the Lync Autoresponce bot and sign out of the Lync client 

		.DESCRIPTION
   			The purpose of Lync-NoBot is to demonstrate how PowerShell can be used to interact with the Lync SDK.
	
		.EXAMPLE
			Lync-Signout
	#>
	
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
		<#
		.SYNOPSIS
			Lync-Shutdown is a PowerShell function to Shutdown the Lync Client 
		
		.DESCRIPTION
   			The purpose of Lync-Shutdown is to demonstrate how PowerShell can be used to interact with the Lync SDK.
	
		.PARAMETER Confirm
				Used to Confirm the Lync Client Shutdown Process

		.EXAMPLE
			Lync-Shutdown -Confirm
	#>
	
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
		<#
		.SYNOPSIS
			Lync-Bot is a PowerShell function to Start the Lync Auto responder bot 
		
		.DESCRIPTION
   			The purpose of Lync-Bot is to demonstrate how PowerShell can be used to interact with the Lync SDK.
	
		.EXAMPLE
			Lync-Bot
	#>
	$ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
		# Register events for current open conversation participants
	foreach ($con in $client.ConversationManager.Conversations)
	{
		# For each participant in the conversation
		$moo = $con.Participants | Where { !$_.IsSelf }
		
		foreach ($mo in $moo)
		{
			try
			{
				if (!(Get-EventSubscriber $mo.Contact.uri))
				{
					Register-ObjectEvent -InputObject $mo.Modalities[1] `
										 -EventName "InstantMessageReceived" `
										 -SourceIdentifier $mo.Contact.uri `
										 -action $action
				}
			}
			catch [system.ArgumentException] { }
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
	$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
	function Global:prompt { [System.String]$Global:usr = $Client.Uri.TrimStart("sip:");"Lync Bot CLI [$usr] [Bot:ON] >" }
}

function Lync-reload
{
	<#
		.SYNOPSIS
			Lync-reload is a PowerShell function to Kill the autoresponce bot and reload the Lync bot module.  
		
		.DESCRIPTION
   			The purpose of Lync-reload is to demonstrate how PowerShell can be used to interact with the Lync SDK.
	
		.EXAMPLE
			Lync-reload
	#>
	
	Lync-NoBot
	Remove-Module 'Lync Bot'
	Import-Module 'Lync Bot'
}

#endregion

#region Runtime INIT
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

#endregion

#region Export Module members
# export module members
Export-ModuleMember lync-State-Change
Export-ModuleMember lync-send-msg
Export-ModuleMember Lync-Bot
Export-ModuleMember Lync-NoBot
Export-ModuleMember Lync-Shutdown
Export-ModuleMember Lync-Signout
Export-ModuleMember Lync-Availability
Export-ModuleMember Lync-reload
Export-ModuleMember do-Rot13
#endregion