########################################################################
# Name: Skype for Business Rate My Call Viewer Tool
# Version: v1.0.5 (16/3/2018)
# Original Release Date: 16/1/2017
# Created By: James Cussen
# Web Site: http://www.myskypelab.com
#
# Purpose: This tool is designed to read Rate My Call data out of the Monitoring database.
# Notes: This is a Powershell tool designed to be run on a machine with the Skype for Business Powershell module installed. To run the tool, open it from the Powershell command line or Right Click and select "Run with Powershell".
# 		 For more information on the requirements for setting up and using this tool please visit http://www.myskypelab.com.
#
# Powershell Version: Supported on Version 3.0 and above
#
# Copyright: Copyright (c) 2018, James Cussen (www.myskypelab.com) All rights reserved.
# Licence: 	Redistribution and use of script, source and binary forms, with or without modification, are permitted provided that the following conditions are met:
#				1) Redistributions of script code must retain the above copyright notice, this list of conditions and the following disclaimer.
#				2) Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
#				3) Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
#				4) This license does not include any resale or commercial use of this software.
#				5) Any portion of this software may not be reproduced, duplicated, copied, sold, resold, or otherwise exploited for any commercial purpose without express written consent of James Cussen.
#			THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; LOSS OF GOODWILL OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
#
# Rate My Call Policy Settings: Set-CSClientPolicy -Identity <PolicyIdentity> -RateMyCallDisplayPercentage 80 -RateMyCallAllowCustomUserFeedback $true
#
# Include Reasons Filter Example: OR Operator would look like this "echo|backgroundnoise"
#
# Prerequsites: 
# - This tool should be run on a machine that has the Skype for Business powershell module installed. This is required because the "Get-CSService" command is used to discover the location of the Montoring Database.
# - The user running the tool needs to have sufficient rights to run select queries on the "QoEMetrics" database and SELECT access on the following tables: Session, AudioStream, CallQualityFeedback, CallQualityFeedbackToken, CallQualityFeedbackTokenDef, User, MediaLine
#
# Release Notes:
# 1.00 Initial Release.
#
# 1.01 Update
#	- Added the ability to select individual Monitoring Servers from a drop down box. This was added for large environments that have multiple Monitoring databases and only want to retrieve statistics from one at a time. By default all monitoring database will be queried.
#	- Added a check for the database version. The tool only works on Skype for Business so the check makes sure the database is at least Version 7 (ie. Skype for Business level).
#
# 1.02 Update
#	- Added date/time localisation checkbox. By default the monitoring server records time is in GMT. This update adds a checkbox to localise all the date/time values to be in the timezone of the server you are running it on (instead of GMT). This changes the date pickers as well as the date displayed in the list and graphs.
#	- Added the ability to zoom in on the Trend Over Time chart. You do this by clicking and dragging the mouse on the area of the graph you want to zoom to, scroll bars will appear so you can scroll the zoomed in view.
#
# 1.03 Update
#	- Fixed an issue with the SQL query used for Video / Audio. The query now gets all records.
#	- Fixed issue with data grid view scroll bar refresh.
#	- Fixed a sorting issue with the Stacked Bar and Trend Over Time Graphs that would cause an issue with the output.
#	- More accurate graphs! When both video and voice are selected the rating data gets listed twice for each call because video calls contain both voice and video ratings. So in previous versions the star ratings were counted as separate calls which artificially inflated the star rating value given. In this version the double counting of this data has been removed from star rating graphs, with the voice and video star rating given by each user being combined. 
#
# 1.04 Update
#	- Now Supports Skype for Business C2R 2016 client Rate My Call issue items. The C2R 2016 client has an entirely new set of rate my call feedback, so the tool has been updated to include these.
#	- Re-worked the graphs again to handle new data
#	- Voice and Video calls don't get listed twice in this version (as it did in the previous version), graph processing was updated from previous version to handle this.
#	- Get Records processing speed was increased by limiting records by date range in SQL query.
#
# 1.05 Update
#	- Total Rows Counter added at the bottom
#	- "Top 10 One Star Users" graph added. This can be used so you can follow up with these users about their bad experiences.
#	- "Top 10 Zero Star Users (Lync 2013 Client)" graph added. This can be used to follow up on Lync 2013 client users that are not responding the Rate My Call dialog.
#
########################################################################



$theVersion = $PSVersionTable.PSVersion
$MajorVersion = $theVersion.Major

Write-Host ""
Write-Host "--------------------------------------------------------------"
Write-Host "Powershell Version Check..." -foreground "yellow"
if($MajorVersion -eq  "1")
{
	Write-Host "This machine only has Version 1 Powershell installed.  This version of Powershell is not supported." -foreground "red"
}
elseif($MajorVersion -eq  "2")
{
	Write-Host "This machine has Version 2 Powershell installed. This version of Powershell is not supported." -foreground "red"
}
elseif($MajorVersion -eq  "3")
{
	Write-Host "This machine has version 3 Powershell installed. CHECK PASSED!" -foreground "green"
}
elseif($MajorVersion -eq  "4")
{
	Write-Host "This machine has version 4 Powershell installed. CHECK PASSED!" -foreground "green"
}
elseif($MajorVersion -eq  "5")
{
	Write-Host "This machine has version 5 Powershell installed. CHECK PASSED!" -foreground "green"
}
else
{
	Write-Host "This machine has version $MajorVersion Powershell installed. Unknown level of support for this version." -foreground "yellow"
}
Write-Host "--------------------------------------------------------------"
Write-Host ""


Function Get-MyModule 
{ 
Param([string]$name) 
	
	if(-not(Get-Module -name $name)) 
	{ 
		if(Get-Module -ListAvailable | Where-Object { $_.name -eq $name }) 
		{ 
			Import-Module -Name $name 
			return $true 
		} #end if module available then import 
		else 
		{ 
			return $false 
		} #module not available 
	} # end if not module 
	else 
	{ 
		return $true 
	} #module already loaded 
} #end function get-MyModule 


$Script:LyncModuleAvailable = $false
$Script:SkypeModuleAvailable = $false

Write-Host "--------------------------------------------------------------"
#Import Lync Module
if(Get-MyModule "Lync")
{
	Invoke-Expression "Import-Module Lync"
	Write-Host "Imported Lync Module..." -foreground "green"
	$Script:LyncModuleAvailable = $true
}
else
{
	Write-Host "Unable to import Lync Module..." -foreground "yellow"
}
#Import SkypeforBusiness Module
if(Get-MyModule "SkypeforBusiness")
{
	Invoke-Expression "Import-Module SkypeforBusiness"
	Write-Host "Imported SkypeforBusiness Module..." -foreground "green"
	$Script:SkypeModuleAvailable = $true
}
else
{
	Write-Host "Unable to import SkypeforBusiness Module... (Expected on a Lync 2013 system)" -foreground "yellow"
}

$scriptpath = $MyInvocation.MyCommand.Path
$dir = Split-Path $scriptpath
$fso = New-Object -ComObject Scripting.FileSystemObject
$shortname = $fso.GetFolder($dir).Path
$ScriptFolder = $shortname


$script:CancelScan = $false

[void][reflection.assembly]::loadwithpartialname("system.drawing")

$TokenDef = @{"1"="DistortedSpeech";"2"="ElectronicFeedback";"3"="BackgroundNoise";"4"="MuffledSpeech";"5"="Echo";"21"="FrozenVideo";"22"="PixelatedVideo";"23"="BlurryImage";"24"="PoorColor";"25"="DarkVideo";"101"="NoSpeechNearSide";"102"="NoSpeechFarSide";"103"="Echo";"104"="IHeardNoise";"105"="VolumeLow";"106"="VoiceCallCutOff";"107"="DistortedSpeech";"108"="TalkingOverEachOther";"201"="NoVideoNearSide";"202"="NoVideoFarSide";"203"="PoorQualityVideo";"204"="FrozenVideo";"205"="VideoCutOff";"206"="DarkVideo";"207"="VideoAudioOutOfSync";}

<#
"101"="I could no hear any sound"
"102"="The other side could not hear any sound"
"103"="I heard echo in the call"
"104"="I heard noise on the call"
"105"="Volume was low"
"106"="Call ened unexpectedly"
"107"="Speech was not natural or sounded distorted"
"108"="We kept interrupting each other"

"101"="NoSpeechNearSide"
"102"="NoSpeechFarSide"
"103"="Echo"
"104"="IHeardNoise"
"105"="VolumeLow"
"106"="VoiceCallCutOff"
"107"="DistortedSpeech"
"108"="TalkingOverEachOther"

"201"="I could not see any video"
"202"="The other side could not see my video"
"203"="Image quality was poor"
"204"="Video Kept Freezing"
"205"="Video stopped unexpectedly"
"206"="The other side was too dark"
"207"="Video was ahead or behind audio"

"201"="NoVideoNearSide"
"202"="NoVideoFarSide"
"203"="PoorQualityVideo"
"204"="FrozenVideo"
"205"="VideoCallCutOff"
"206"="DarkVideo"
"207"="VideoAudioOutOfSync"
#>

# Set up the form  ============================================================

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "Skype4B Rate My Call Viewer Tool v1.05"
$objForm.Size = New-Object System.Drawing.Size(910,570) 
$objForm.MinimumSize = New-Object System.Drawing.Size(860,340) 
$objForm.StartPosition = "CenterScreen"
[byte[]]$WindowIcon = @(137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 32, 0, 0, 0, 32, 8, 6, 0, 0, 0, 115, 122, 122, 244, 0, 0, 0, 6, 98, 75, 71, 68, 0, 255, 0, 255, 0, 255, 160, 189, 167, 147, 0, 0, 0, 9, 112, 72, 89, 115, 0, 0, 11, 19, 0, 0, 11, 19, 1, 0, 154, 156, 24, 0, 0, 0, 7, 116, 73, 77, 69, 7, 225, 7, 26, 1, 36, 51, 211, 178, 227, 235, 0, 0, 5, 235, 73, 68, 65, 84, 88, 195, 197, 151, 91, 108, 92, 213, 21, 134, 191, 189, 207, 57, 115, 159, 216, 78, 176, 27, 98, 72, 226, 88, 110, 66, 66, 34, 185, 161, 168, 193, 73, 21, 17, 2, 2, 139, 75, 164, 182, 106, 145, 170, 190, 84, 74, 104, 65, 16, 144, 218, 138, 138, 244, 173, 69, 106, 101, 42, 129, 42, 149, 170, 162, 15, 168, 168, 151, 7, 4, 22, 180, 1, 41, 92, 172, 52, 196, 68, 105, 130, 19, 138, 98, 76, 154, 27, 174, 227, 248, 58, 247, 57, 103, 175, 62, 236, 241, 177, 199, 246, 140, 67, 26, 169, 251, 237, 236, 61, 179, 215, 191, 214, 191, 214, 191, 214, 86, 188, 62, 37, 252, 31, 151, 174, 123, 42, 224, 42, 72, 56, 138, 152, 99, 191, 175, 247, 114, 107, 29, 172, 75, 106, 94, 254, 74, 156, 109, 13, 58, 180, 155, 53, 240, 216, 64, 129, 63, 156, 43, 95, 55, 0, 106, 62, 5, 158, 134, 83, 59, 147, 116, 36, 106, 7, 103, 188, 44, 228, 13, 120, 202, 126, 151, 12, 100, 3, 225, 183, 231, 203, 60, 55, 88, 66, 4, 80, 215, 0, 96, 89, 68, 113, 97, 87, 138, 180, 3, 163, 101, 120, 116, 160, 192, 161, 81, 159, 203, 69, 33, 230, 40, 58, 27, 52, 251, 215, 69, 248, 198, 74, 183, 238, 165, 175, 141, 248, 60, 114, 178, 192, 165, 188, 44, 9, 100, 22, 128, 192, 127, 238, 73, 209, 18, 81, 252, 109, 52, 224, 222, 247, 179, 179, 46, 206, 93, 102, 142, 119, 193, 76, 216, 96, 247, 13, 46, 223, 189, 201, 101, 207, 74, 143, 148, 99, 183, 159, 250, 184, 72, 207, 96, 169, 46, 136, 16, 192, 183, 91, 61, 94, 233, 140, 241, 81, 198, 176, 229, 173, 204, 226, 198, 175, 102, 5, 194, 243, 157, 113, 246, 221, 236, 225, 42, 232, 29, 9, 184, 255, 104, 174, 62, 0, 165, 192, 239, 78, 163, 129, 174, 195, 57, 14, 143, 5, 255, 115, 114, 197, 29, 197, 200, 221, 41, 82, 14, 188, 63, 30, 240, 245, 190, 220, 162, 145, 208, 0, 141, 174, 66, 1, 37, 129, 195, 163, 254, 34, 40, 1, 191, 70, 25, 250, 50, 75, 197, 156, 149, 15, 132, 27, 254, 62, 205, 229, 178, 176, 163, 201, 161, 103, 115, 172, 182, 14, 196, 181, 53, 114, 38, 107, 64, 22, 194, 92, 147, 80, 200, 67, 105, 50, 247, 165, 171, 156, 104, 141, 105, 70, 186, 211, 200, 131, 105, 214, 46, 82, 53, 69, 3, 119, 244, 217, 240, 63, 177, 214, 35, 233, 170, 250, 66, 164, 20, 11, 221, 52, 240, 171, 77, 49, 114, 6, 198, 74, 18, 158, 106, 5, 239, 110, 79, 208, 236, 41, 254, 93, 16, 206, 102, 204, 162, 30, 14, 78, 27, 158, 60, 93, 68, 1, 7, 191, 150, 176, 73, 60, 31, 64, 182, 178, 185, 49, 169, 103, 80, 132, 235, 166, 164, 38, 238, 64, 66, 67, 104, 94, 224, 229, 206, 56, 111, 93, 182, 116, 61, 246, 81, 177, 118, 166, 107, 248, 253, 121, 43, 92, 119, 52, 106, 86, 39, 245, 66, 0, 147, 101, 9, 105, 188, 171, 165, 186, 198, 127, 179, 57, 202, 233, 233, 106, 216, 9, 79, 113, 169, 96, 216, 119, 179, 135, 47, 112, 240, 114, 185, 110, 169, 77, 149, 132, 95, 159, 181, 32, 182, 54, 58, 139, 83, 112, 231, 7, 121, 0, 126, 210, 17, 129, 96, 150, 134, 213, 9, 205, 84, 185, 42, 29, 121, 103, 91, 130, 15, 38, 45, 228, 105, 95, 40, 207, 97, 173, 209, 83, 124, 179, 213, 227, 153, 13, 81, 16, 91, 205, 247, 174, 116, 113, 42, 118, 31, 89, 227, 86, 37, 109, 8, 224, 189, 97, 159, 178, 64, 71, 82, 207, 166, 129, 192, 75, 231, 203, 180, 68, 170, 235, 252, 95, 57, 195, 150, 138, 218, 156, 43, 8, 70, 102, 43, 98, 96, 103, 146, 63, 119, 198, 120, 115, 216, 210, 243, 179, 245, 81, 222, 248, 106, 156, 141, 73, 77, 201, 192, 109, 141, 14, 86, 171, 231, 39, 161, 99, 209, 158, 43, 152, 48, 156, 237, 41, 205, 123, 163, 1, 174, 99, 55, 38, 3, 225, 209, 142, 40, 7, 78, 23, 217, 182, 220, 2, 120, 247, 202, 172, 59, 27, 155, 28, 90, 163, 138, 76, 32, 28, 159, 12, 192, 23, 30, 110, 181, 148, 238, 63, 85, 64, 128, 166, 121, 149, 160, 23, 118, 96, 21, 122, 255, 226, 150, 40, 103, 178, 134, 132, 182, 123, 167, 50, 134, 95, 222, 18, 229, 108, 198, 112, 99, 212, 238, 29, 155, 156, 5, 240, 253, 53, 54, 84, 127, 25, 246, 9, 4, 214, 175, 112, 104, 139, 107, 46, 20, 132, 129, 41, 179, 196, 60, 96, 108, 228, 155, 61, 107, 60, 237, 41, 140, 82, 100, 138, 66, 186, 146, 151, 67, 89, 195, 119, 142, 231, 65, 36, 212, 251, 209, 188, 132, 212, 116, 85, 18, 236, 233, 143, 139, 0, 252, 174, 34, 62, 71, 39, 131, 80, 107, 138, 82, 11, 128, 182, 213, 176, 33, 169, 33, 128, 159, 174, 143, 176, 231, 104, 30, 20, 172, 170, 120, 187, 111, 181, 199, 171, 151, 124, 80, 48, 94, 17, 204, 111, 173, 246, 160, 44, 188, 182, 45, 73, 103, 131, 189, 110, 120, 218, 240, 192, 74, 151, 29, 77, 22, 80, 207, 80, 137, 6, 79, 227, 42, 136, 42, 112, 230, 244, 153, 16, 128, 18, 155, 193, 0, 127, 237, 74, 48, 81, 18, 50, 190, 128, 8, 55, 198, 236, 207, 186, 251, 243, 161, 10, 205, 112, 255, 189, 85, 46, 178, 103, 25, 61, 67, 37, 222, 24, 177, 168, 142, 237, 74, 209, 28, 213, 76, 248, 66, 206, 192, 67, 95, 242, 56, 240, 229, 8, 253, 21, 26, 126, 176, 54, 178, 112, 34, 18, 5, 63, 255, 180, 196, 211, 237, 17, 20, 240, 236, 39, 37, 11, 79, 89, 158, 247, 159, 242, 57, 50, 211, 164, 20, 60, 126, 178, 64, 68, 131, 163, 96, 239, 201, 2, 34, 112, 100, 220, 231, 135, 107, 35, 188, 114, 209, 103, 119, 179, 67, 163, 171, 24, 200, 24, 122, 134, 138, 124, 158, 23, 86, 197, 53, 23, 239, 74, 242, 112, 171, 199, 243, 131, 69, 112, 212, 188, 137, 40, 0, 121, 48, 109, 109, 244, 102, 174, 105, 8, 92, 151, 208, 244, 109, 79, 112, 177, 32, 220, 182, 76, 115, 123, 95, 142, 254, 137, 32, 188, 127, 172, 59, 133, 163, 160, 225, 245, 105, 112, 213, 188, 42, 112, 224, 197, 138, 108, 158, 216, 153, 248, 226, 61, 88, 224, 79, 91, 227, 180, 189, 157, 97, 115, 74, 115, 104, 44, 160, 127, 78, 153, 162, 160, 28, 64, 84, 171, 218, 101, 184, 247, 159, 5, 174, 248, 176, 37, 165, 121, 118, 83, 244, 11, 5, 161, 179, 209, 225, 76, 222, 240, 194, 230, 24, 142, 134, 61, 253, 121, 112, 170, 69, 172, 33, 162, 24, 47, 75, 157, 177, 92, 65, 87, 95, 22, 128, 31, 183, 69, 56, 176, 33, 90, 37, 205, 245, 214, 241, 241, 128, 67, 35, 1, 39, 38, 13, 94, 239, 52, 147, 229, 234, 255, 221, 211, 234, 17, 85, 208, 119, 37, 176, 237, 116, 177, 169, 120, 38, 148, 91, 151, 59, 124, 216, 149, 168, 12, 153, 1, 123, 79, 228, 25, 206, 203, 82, 47, 137, 186, 244, 100, 187, 211, 36, 52, 220, 255, 97, 158, 222, 138, 84, 235, 26, 131, 26, 199, 198, 3, 154, 14, 102, 152, 240, 133, 7, 90, 28, 62, 223, 157, 226, 165, 173, 113, 86, 120, 138, 168, 14, 29, 176, 169, 163, 150, 54, 254, 199, 219, 227, 36, 52, 156, 206, 25, 122, 47, 148, 107, 191, 11, 22, 72, 165, 130, 95, 108, 140, 241, 163, 54, 111, 230, 46, 138, 6, 2, 17, 130, 202, 212, 173, 21, 228, 12, 220, 249, 143, 28, 3, 19, 166, 170, 53, 183, 196, 20, 71, 182, 39, 105, 139, 219, 205, 230, 131, 25, 70, 75, 114, 245, 0, 102, 100, 122, 69, 76, 177, 171, 217, 229, 153, 142, 8, 183, 166, 106, 243, 112, 46, 47, 97, 146, 165, 92, 104, 175, 140, 106, 99, 62, 108, 122, 39, 195, 112, 65, 234, 191, 140, 150, 10, 37, 70, 64, 43, 54, 164, 53, 77, 17, 133, 8, 92, 42, 26, 118, 44, 119, 121, 170, 61, 66, 103, 186, 26, 220, 80, 78, 120, 238, 179, 18, 47, 12, 150, 170, 43, 226, 154, 0, 92, 197, 155, 0, 20, 237, 203, 172, 238, 127, 50, 101, 108, 239, 175, 147, 36, 238, 117, 125, 234, 86, 12, 125, 58, 51, 100, 106, 150, 124, 36, 254, 23, 153, 41, 93, 205, 81, 212, 105, 60, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130)
$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
$objForm.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
$objForm.KeyPreview = $True
$objForm.TabStop = $false


$MyLinkLabel = New-Object System.Windows.Forms.LinkLabel
$MyLinkLabel.Location = New-Object System.Drawing.Size(690,5)
$MyLinkLabel.Size = New-Object System.Drawing.Size(180,15)
$MyLinkLabel.DisabledLinkColor = [System.Drawing.Color]::Red
$MyLinkLabel.VisitedLinkColor = [System.Drawing.Color]::Blue
$MyLinkLabel.LinkBehavior = [System.Windows.Forms.LinkBehavior]::HoverUnderline
$MyLinkLabel.LinkColor = [System.Drawing.Color]::Navy
$MyLinkLabel.TabStop = $False
$MyLinkLabel.Text = "Created by: www.myskypelab.com"
$MyLinkLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Top
$MyLinkLabel.add_click(
{
	 [system.Diagnostics.Process]::start("http://www.myskypelab.com")
})
$objForm.Controls.Add($MyLinkLabel)



#$Network Label ============================================================
$StartTimeLabel = New-Object System.Windows.Forms.Label
$StartTimeLabel.Location = New-Object System.Drawing.Size(30,30) 
$StartTimeLabel.Size = New-Object System.Drawing.Size(35,15) 
$StartTimeLabel.Text = "Start: "
$StartTimeLabel.TabStop = $false
$StartTimeLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$objForm.Controls.Add($StartTimeLabel)


$StartTimePicker = New-Object System.Windows.Forms.DateTimePicker
$StartTimePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
$StartTimePicker.CustomFormat = "yy/MM/dd HH:mm";
$StartTimePicker.ShowUpDown = $true
$StartTimePicker.Location = New-Object Drawing.Size(65,27)
$StartTimePicker.Width = 110
$StartTimePicker.tabIndex = 1
$StartTimePicker.Value = ((Get-Date) - (New-TimeSpan -days 30)).ToUniversalTime()
$objForm.Controls.Add($StartTimePicker)

#$Network Label ============================================================
$EndTimeLabel = New-Object System.Windows.Forms.Label
$EndTimeLabel.Location = New-Object System.Drawing.Size(35,57) 
$EndTimeLabel.Size = New-Object System.Drawing.Size(30,15) 
$EndTimeLabel.Text = "End: "
$EndTimeLabel.TabStop = $false
$EndTimeLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$objForm.Controls.Add($EndTimeLabel)


$EndTimePicker = New-Object System.Windows.Forms.DateTimePicker
$EndTimePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
$EndTimePicker.CustomFormat = "yy/MM/dd HH:mm";
$EndTimePicker.ShowUpDown = $true
$EndTimePicker.Location = New-Object Drawing.Size(65,53)
$EndTimePicker.Width = 110
$EndTimePicker.tabIndex = 2
$EndTimePicker.Value = ((Get-Date)).ToUniversalTime()
$objForm.Controls.Add($EndTimePicker)



#TimeFormatCheckBox ============================================================
$TimeFormatCheckBox = New-Object System.Windows.Forms.Checkbox 
$TimeFormatCheckBox.Location = New-Object System.Drawing.Size(182,33) 
$TimeFormatCheckBox.Size = New-Object System.Drawing.Size(20,20)
$TimeFormatCheckBox.TabIndex = 4
$TimeFormatCheckBox.Add_Click({
	#Does nothing
	if($TimeFormatCheckBox.Checked)
	{
		$TimeFormatLabel.Text = "GMT"
		$StartTimePicker.Value = (($StartTimePicker.Value)).ToUniversalTime()
		$EndTimePicker.Value = (($EndTimePicker.Value)).ToUniversalTime()
		$titleColumn0.HeaderText = "Date/Time (GMT)"
		$dgv.Rows.Clear()
	}
	else
	{
		$TimeFormatLabel.Text = "Local"
		$StartTimePicker.Value = [System.TimeZone]::CurrentTimeZone.ToLocalTime($StartTimePicker.Value)
		$EndTimePicker.Value = [System.TimeZone]::CurrentTimeZone.ToLocalTime($EndTimePicker.Value)
		$titleColumn0.HeaderText = "Date/Time (Local)"
		$dgv.Rows.Clear()
	}
})
$objForm.Controls.Add($TimeFormatCheckBox) 
$TimeFormatCheckBox.Checked = $true


#TimeFormat Label ============================================================
$TimeFormatLabel = New-Object System.Windows.Forms.Label
$TimeFormatLabel.Location = New-Object System.Drawing.Size(175,55) 
$TimeFormatLabel.Size = New-Object System.Drawing.Size(35,15) 
$TimeFormatLabel.Text = "GMT"
$TimeFormatLabel.TabStop = $false
$TimeFormatLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$objForm.Controls.Add($TimeFormatLabel)



#Rating Label ============================================================
$RatingLabel = New-Object System.Windows.Forms.Label
$RatingLabel.Location = New-Object System.Drawing.Size(220,32) 
$RatingLabel.Size = New-Object System.Drawing.Size(40,15) 
$RatingLabel.Text = "Rating: "
$RatingLabel.TabStop = $false
$RatingLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$objForm.Controls.Add($RatingLabel)



#Rating Drop Down box ============================================================
$RatingTextBox = New-Object System.Windows.Forms.ComboBox 
$RatingTextBox.Location = New-Object System.Drawing.Size(265,30) 
$RatingTextBox.Size = New-Object System.Drawing.Size(80,15) 
$RatingTextBox.DropDownHeight = 100 
$RatingTextBox.DropDownWidth = 60 
$RatingTextBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$RatingTextBox.tabIndex = 3
$RatingTextBox.Sorted = $true
$RatingTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$objForm.Controls.Add($RatingTextBox) 

[void] $RatingTextBox.Items.Add("0")
[void] $RatingTextBox.Items.Add("1")
[void] $RatingTextBox.Items.Add("2")
[void] $RatingTextBox.Items.Add("3")
[void] $RatingTextBox.Items.Add("4")
[void] $RatingTextBox.Items.Add("5")
[void] $RatingTextBox.Items.Add("ALL")

$RatingTextBox.SelectedIndex = $RatingTextBox.FindStringExact("ALL")


#And Above Label ============================================================
$AndAboveLabel = New-Object System.Windows.Forms.Label
$AndAboveLabel.Location = New-Object System.Drawing.Size(220,55) 
$AndAboveLabel.Size = New-Object System.Drawing.Size(43,15) 
$AndAboveLabel.Text = "Above:"
$AndAboveLabel.TabStop = $false
$AndAboveLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$AndAboveLabel.Add_Click({

	if($AndAboveCheckBox.Checked)
	{
		$AndAboveCheckBox.Checked = $false
	}
	else
	{
		$AndAboveCheckBox.Checked = $true
		$AndBelowCheckBox.Checked = $false
	}
})
$objForm.Controls.Add($AndAboveLabel)


#AndAboveCheckBox ============================================================
$AndAboveCheckBox = New-Object System.Windows.Forms.Checkbox 
$AndAboveCheckBox.Location = New-Object System.Drawing.Size(265,53) 
$AndAboveCheckBox.Size = New-Object System.Drawing.Size(20,20)
$AndAboveCheckBox.TabIndex = 4
$AndAboveCheckBox.Add_Click({
	#Does nothing
	
	if($AndAboveCheckBox.Checked)
	{
		$AndBelowCheckBox.Checked = $false
	}
})
$objForm.Controls.Add($AndAboveCheckBox) 
$AndAboveCheckBox.Checked = $false



#And Below Label ============================================================
$AndBelowLabel = New-Object System.Windows.Forms.Label
$AndBelowLabel.Location = New-Object System.Drawing.Size(288,55) 
$AndBelowLabel.Size = New-Object System.Drawing.Size(38,15) 
$AndBelowLabel.Text = "Below:"
$AndBelowLabel.TabStop = $false
$AndBelowLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$AndBelowLabel.Add_Click({

	if($AndBelowCheckBox.Checked)
	{
		$AndBelowCheckBox.Checked = $false
	}
	else
	{
		$AndBelowCheckBox.Checked = $true
		$AndAboveCheckBox.Checked = $false
	}
})
$objForm.Controls.Add($AndBelowLabel)

#AndBelowCheckBox ============================================================
$AndBelowCheckBox = New-Object System.Windows.Forms.Checkbox 
$AndBelowCheckBox.Location = New-Object System.Drawing.Size(330,53) 
$AndBelowCheckBox.Size = New-Object System.Drawing.Size(20,20)
$AndBelowCheckBox.TabIndex = 4
$AndBelowCheckBox.Add_Click({
	#Does nothing
	if($AndBelowCheckBox.Checked)
	{
		$AndAboveCheckBox.Checked = $false
	}
})
$objForm.Controls.Add($AndBelowCheckBox) 
$AndAboveCheckBox.Checked = $false



#EventIDLabel Label ============================================================
$EventIDLabel2 = New-Object System.Windows.Forms.Label
$EventIDLabel2.Location = New-Object System.Drawing.Size(355,50) 
$EventIDLabel2.Size = New-Object System.Drawing.Size(45,15) 
$EventIDLabel2.Text = "Include"
$EventIDLabel2.TabStop = $false
$EventIDLabel2.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$objForm.Controls.Add($EventIDLabel2)

#EventIDLabel Label ============================================================
$EventIDLabel = New-Object System.Windows.Forms.Label
$EventIDLabel.Location = New-Object System.Drawing.Size(355,65) 
$EventIDLabel.Size = New-Object System.Drawing.Size(52,15) 
$EventIDLabel.Text = "Reasons: "
$EventIDLabel.TabStop = $false
$EventIDLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$objForm.Controls.Add($EventIDLabel)



#$Reason Text box ============================================================
$ReasonTextBox = New-Object System.Windows.Forms.TextBox 
$ReasonTextBox.Location = New-Object System.Drawing.Size(410,55) 
$ReasonTextBox.Size = New-Object System.Drawing.Size(180,15) 
$ReasonTextBox.tabIndex = 6
$ReasonTextBox.Text = ""
$ReasonTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$objForm.Controls.Add($ReasonTextBox) 


#User Label ============================================================
$UserLabel = New-Object System.Windows.Forms.Label
$UserLabel.Location = New-Object System.Drawing.Size(355,30) 
$UserLabel.Size = New-Object System.Drawing.Size(52,15) 
$UserLabel.Text = "SIP URI: "
$UserLabel.TabStop = $false
$UserLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$objForm.Controls.Add($UserLabel)


#$UserTextBox ============================================================
$UserTextBox = New-Object System.Windows.Forms.TextBox 
$UserTextBox.Location = New-Object System.Drawing.Size(410,30) 
$UserTextBox.Size = New-Object System.Drawing.Size(180,15) 
$UserTextBox.tabIndex = 5
$UserTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$objForm.Controls.Add($UserTextBox)



#Find Previous button
$FindPreviousButton = New-Object System.Windows.Forms.Button
$FindPreviousButton.Location = New-Object System.Drawing.Size(30,85)
$FindPreviousButton.Size = New-Object System.Drawing.Size(120,20)
$FindPreviousButton.Text = "<- Find Previous"
$FindPreviousButton.tabIndex = 10
$FindPreviousButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$FindPreviousButton.Add_Click(
{
	$StatusLabel.Text = ""
	[string]$FindString = $FindNextTextBox.Text
	$selectedRowCount = $dgv.SelectedRows[0].Index
	$selectedRowCount = $selectedRowCount - 1
	$WasFound = $false
	for($i=$selectedRowCount; $i -ge 0; $i--)
	{
		for($j=1; $j -lt 5; $j++)
		{
			$cellstr = $dgv.Rows[$i].Cells[$j].Value.ToString()
			
			if($cellstr -ne $null)
			{
				if($cellstr -match $FindString)
				{
					$WasFound = $true
				}
			}
		}
		if($WasFound)
		{
			$dgv.Rows[$i].Selected = $true
			$dgv.FirstDisplayedScrollingRowIndex = $i
			Write-Host "Found Search String at row: $i" -foreground "green"
			break
		}
	}
	if(!$WasFound)
	{
		Write-Host "Search string was not found" -foreground "red"
		$StatusLabel.Text = "Search string was not found"
	}
	[System.Windows.Forms.Application]::DoEvents()
})
$objForm.Controls.Add($FindPreviousButton)


#FindNext Text box ============================================================
$FindNextTextBox = New-Object System.Windows.Forms.TextBox 
$FindNextTextBox.Location = New-Object System.Drawing.Size(160,85) 
$FindNextTextBox.Size = New-Object System.Drawing.Size(120,15) 
$FindNextTextBox.tabIndex = 11
$FindNextTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$objForm.Controls.Add($FindNextTextBox)
$FindNextTextBox.add_KeyUp(
{
	if ($_.KeyCode -eq "Enter") 
	{
		#FIND NEXT
		$StatusLabel.Text = ""
		[string]$FindString = $FindNextTextBox.Text
		$selectedRowCount = $dgv.SelectedRows[0].Index
		$selectedRowCount = $selectedRowCount + 1
		$WasFound = $false
		for($i=$selectedRowCount; $i -lt $dgv.Rows.Count; $i++)
		{
			for($j=1; $j -lt 5; $j++)
			{
				if($dgv.Rows[$i].Cells[$j].Value -ne $null)
				{
					$cellstr = $dgv.Rows[$i].Cells[$j].Value.ToString()
				}
				
				if($cellstr -ne $null)
				{
					if($cellstr -match $FindString)
					{
						$WasFound = $true
					}
				}
			}
			if($WasFound)
			{
				$dgv.Rows[$i].Selected = $true
				$dgv.FirstDisplayedScrollingRowIndex = $i
				Write-Host "Found Search String at row: $i" -foreground "green"
				break
			}
		}
		if(!$WasFound)
		{
			Write-Host "Search string was not found" -foreground "red"
			$StatusLabel.Text = "Search string was not found"
		}
		[System.Windows.Forms.Application]::DoEvents()
	}
})


#Find Next button
$FindNextButton = New-Object System.Windows.Forms.Button
$FindNextButton.Location = New-Object System.Drawing.Size(290,85)
$FindNextButton.Size = New-Object System.Drawing.Size(120,20)
$FindNextButton.Text = "Find Next ->"
$FindNextButton.tabIndex = 12
$FindNextButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$FindNextButton.Add_Click(
{
	$StatusLabel.Text = ""
	[string]$FindString = $FindNextTextBox.Text
	$selectedRowCount = $dgv.SelectedRows[0].Index
	$selectedRowCount = $selectedRowCount + 1
	$WasFound = $false
	for($i=$selectedRowCount; $i -lt $dgv.Rows.Count; $i++)
	{
		for($j=1; $j -lt 5; $j++)
		{
			if($dgv.Rows[$i].Cells[$j].Value -ne $null)
			{
				$cellstr = $dgv.Rows[$i].Cells[$j].Value.ToString()
			}
			
			if($cellstr -ne $null)
			{
				if($cellstr -match $FindString)
				{
					$WasFound = $true
				}
			}
		}
		if($WasFound)
		{
			$dgv.Rows[$i].Selected = $true
			$dgv.FirstDisplayedScrollingRowIndex = $i
			Write-Host "Found Search String at row: $i" -foreground "green"
			break
		}
	}
	if(!$WasFound)
	{
		Write-Host "Search string was not found" -foreground "red"
		$StatusLabel.Text = "Search string was not found"
	}
	[System.Windows.Forms.Application]::DoEvents()
})
$objForm.Controls.Add($FindNextButton)



#Monitoring Database Label ============================================================
$MonitoringDatabaseLabel = New-Object System.Windows.Forms.Label
$MonitoringDatabaseLabel.Location = New-Object System.Drawing.Size(533,88) 
$MonitoringDatabaseLabel.Size = New-Object System.Drawing.Size(112,15) 
$MonitoringDatabaseLabel.Text = "Monitoring Database:"
$MonitoringDatabaseLabel.TabStop = $false
$MonitoringDatabaseLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Top
$objForm.Controls.Add($MonitoringDatabaseLabel)



#Monitoring Server Drop Down box ============================================================
$MonitoringServerDropDown = New-Object System.Windows.Forms.ComboBox 
$MonitoringServerDropDown.Location = New-Object System.Drawing.Size(645,85) 
$MonitoringServerDropDown.Size = New-Object System.Drawing.Size(210,15) 
$MonitoringServerDropDown.DropDownHeight = 100 
$MonitoringServerDropDown.DropDownWidth = 220 
$MonitoringServerDropDown.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$MonitoringServerDropDown.tabIndex = 3
$MonitoringServerDropDown.Sorted = $true
$MonitoringServerDropDown.Anchor = [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Top
$objForm.Controls.Add($MonitoringServerDropDown) 

# Monitoring Servers
[void] $MonitoringServerDropDown.Items.Add("ALL")
Get-CSService -MonitoringDatabase | select-object PoolFQDN | ForEach-Object {[void] $MonitoringServerDropDown.Items.Add($_.PoolFQDN)}

$MonitoringServerDropDown.SelectedItem = "ALL"



#VoiceCheckBox ============================================================
$VoiceCheckBox = New-Object System.Windows.Forms.Checkbox 
$VoiceCheckBox.Location = New-Object System.Drawing.Size(615,30) 
$VoiceCheckBox.Size = New-Object System.Drawing.Size(20,20)
$VoiceCheckBox.TabIndex = 7
$VoiceCheckBox.Add_Click({
	#Does nothing
})
$objForm.Controls.Add($VoiceCheckBox) 
$VoiceCheckBox.Checked = $true

#VoiceLabel ============================================================
$VoiceLabel = New-Object System.Windows.Forms.Label
$VoiceLabel.Location = New-Object System.Drawing.Size(635,32) 
$VoiceLabel.Size = New-Object System.Drawing.Size(60,15) 
$VoiceLabel.Text = "Voice"
$VoiceLabel.TabStop = $false
$VoiceLabel.Add_Click({
	if($VoiceCheckBox.Checked)
	{
		$VoiceCheckBox.Checked = $false
	}
	else
	{
		$VoiceCheckBox.Checked = $true
	}
})
$objForm.Controls.Add($VoiceLabel)

#VideoCheckBox ============================================================
$VideoCheckBox = New-Object System.Windows.Forms.Checkbox 
$VideoCheckBox.Location = New-Object System.Drawing.Size(615,55) 
$VideoCheckBox.Size = New-Object System.Drawing.Size(20,20)
$VideoCheckBox.TabIndex = 8
$VideoCheckBox.Add_Click({
	#Does nothing
})
$objForm.Controls.Add($VideoCheckBox) 
$VideoCheckBox.Checked = $true

#Video Label ============================================================
$VideoLabel = New-Object System.Windows.Forms.Label
$VideoLabel.Location = New-Object System.Drawing.Size(635,57) 
$VideoLabel.Size = New-Object System.Drawing.Size(60,15) 
$VideoLabel.Text = "Video"
$VideoLabel.TabStop = $false
$VideoLabel.Add_Click({
	if($VideoCheckBox.Checked)
	{
		$VideoCheckBox.Checked = $false
	}
	else
	{
		$VideoCheckBox.Checked = $true
	}
})
$objForm.Controls.Add($VideoLabel)


$CancelDiscoverButton = New-Object System.Windows.Forms.Button
$CancelDiscoverButton.Location = New-Object System.Drawing.Size(700,38)
$CancelDiscoverButton.Size = New-Object System.Drawing.Size(120,25)
$CancelDiscoverButton.Text = "CANCEL SCAN..."
$CancelDiscoverButton.ForeColor = "red"
$CancelDiscoverButton.Visible = $false
$CancelDiscoverButton.Add_Click(
{
	$script:CancelScan = $true
	$CancelDiscoverButton.Enabled = $false
	$FilterButton.Enabled = $true
	$FindPreviousButton.Enabled = $true
	$FindNextButton.Enabled = $true
	$ChartButton.Enabled = $true
	$ExportButton.Enabled = $true
	$dgv.Enabled = $true
	[System.Windows.Forms.Application]::DoEvents()
}
)
$objForm.Controls.Add($CancelDiscoverButton)


#Filter button ============================================================
$FilterButton = New-Object System.Windows.Forms.Button
$FilterButton.Location = New-Object System.Drawing.Size(700,38)
$FilterButton.Size = New-Object System.Drawing.Size(120,25)
$FilterButton.Text = "Get Ratings"
$FilterButton.tabIndex = 9
$FilterButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$FilterButton.Add_Click(
{
	$StatusLabel.Text = "Please Wait..."
	$FilterButton.Text = "Getting Logs"
	$FilterButton.Enabled = $false
	$FindPreviousButton.Enabled = $false
	$FindNextButton.Enabled = $false
	$ExportButton.Enabled = $false
	$ChartButton.Enabled = $false
	$dgv.Enabled = $false
	$CancelDiscoverButton.Visible = $true
	[System.Windows.Forms.Application]::DoEvents()
	
	Get-CallQualityEvents
	
	$FilterButton.Enabled = $true
	$FindPreviousButton.Enabled = $true
	$FindNextButton.Enabled = $true
	$ExportButton.Enabled = $true
	$ChartButton.Enabled = $true
	$dgv.Enabled = $true
	$FilterButton.Text = "Get Ratings"
	$StatusLabel.Text = ""
	
	if($dgv.Rows.Count > 0)
	{
		$dgv.Rows[0].Selected = $true
		$dgv.FirstDisplayedScrollingRowIndex = $i
	}
	
	$CancelDiscoverButton.Visible = $false
		
	[System.Windows.Forms.Application]::DoEvents()
})
$objForm.Controls.Add($FilterButton)


# Add a groupbox ============================================================
$GroupsBox = New-Object System.Windows.Forms.Groupbox
$GroupsBox.Location = New-Object System.Drawing.Size(25,15) 
$GroupsBox.Size = New-Object System.Drawing.Size(850,67) 
$GroupsBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Left
$GroupsBox.TabStop = $False
$objForm.Controls.Add($GroupsBox)



#Chart button ============================================================
$ChartButton = New-Object System.Windows.Forms.Button
$ChartButton.Location = New-Object System.Drawing.Size(250,480)
$ChartButton.Size = New-Object System.Drawing.Size(150,25)
$ChartButton.Text = "Graphs"
$ChartButton.tabIndex = 13
$ChartButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$ChartButton.Add_Click(
{
	$StatusLabel.Text = "Graphing..."
	
	$ChartDialogReturn = ChartDialog -Message "Graphs!" -WindowTitle "Graphs!" -DefaultText "Graphs!"
	
	$StatusLabel.Text = ""
})
$objForm.Controls.Add($ChartButton)




#File Export Browse button ============================================================
$ExportButton = New-Object System.Windows.Forms.Button
$ExportButton.Location = New-Object System.Drawing.Size(420,480)
$ExportButton.Size = New-Object System.Drawing.Size(150,25)
$ExportButton.Text = "Export CSV"
$ExportButton.tabIndex = 14
$ExportButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$ExportButton.Add_Click(
{
	$filename = ""
	
	$StatusLabel.Text = "Please Wait... Exporting"
	$FilterButton.Enabled = $false
	$dgv.Enabled = $false
	$ExportButton.Enabled = $false
	$FilterButton.Enabled = $false
	$ChartButton.Enabled = $false
	$FindPreviousButton.Enabled = $false
	$FindNextButton.Enabled = $false
	[System.Windows.Forms.Application]::DoEvents()
	
	#Get the start and end date
	$StartDate = $StartTimePicker.Value.ToString("yyyy.MM.dd_HH.mm")
	$EndDate = $EndTimePicker.Value.ToString("yyyy.MM.dd_HH.mm")
	
	
	#File Dialog
	[string] $pathVar = "C:\"
	$Filter="All Files (*.*)|*.*"
	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
	$objDialog = New-Object System.Windows.Forms.SaveFileDialog
	#$objDialog.InitialDirectory = 
	$objDialog.FileName = "CallRatingExport_${StartDate}-${EndDate}.csv"
	$objDialog.Filter = $Filter
	$objDialog.Title = "Export File Name"
	$objDialog.CheckFileExists = $false
	$Show = $objDialog.ShowDialog()
	if ($Show -eq "OK")
	{
		[string]$content = ""
		$filename = $objDialog.FileName
	}
	else
	{
		$FilterButton.Enabled = $true
		$dgv.Enabled = $true
		$StatusLabel.Text = ""
		$ExportButton.Enabled = $true
		$FilterButton.Enabled = $true
		$FindPreviousButton.Enabled = $true
		$ChartButton.Enabled = $true
		$FindNextButton.Enabled = $true
		[System.Windows.Forms.Application]::DoEvents()
		return
	}
	
	if($filename -ne "" -and $filename -ne $null)
	{
		$csv = "`"ConferenceDateTime`",`"Caller`",`"Type`",`"Rating`",`"Reason(s)`",`"Feedback Text`"`r`n"
		$CurrentRowCount = 0;
		$TotalRows = $dgv.Rows.Count;
		for($i=0; $i -lt $dgv.Rows.Count; $i++)
		{
			$fileContent = ""
			for($j=0; $j -lt 6; $j++)
			{
				$cellstr = $dgv.Rows[$i].Cells[$j].Value
				
				if($cellstr -ne $null)
				{
					[string]$cellString = $cellstr.ToString()
					$fileContent += "`"$cellString`""
					
					if($j -le 6)
					{
						$fileContent += ","
					}
				}
				else
				{
					$fileContent += ","
				}
			}
			$csv += "$fileContent`r`n"
			$CurrentRowCount++
			if($CurrentRowCount % 100 -eq 0)
			{
				$Percentage = [math]::round(($CurrentRowCount / $TotalRows) * 100, 0)
				$StatusLabel.Text = "Please Wait... Exporting (${Percentage}%)"
				[System.Windows.Forms.Application]::DoEvents()
			}
			
		}
		Write-host "Writing File... $filename" -foreground "green"
		#Excel seems to only like UTF-8 for CSV files...
		$csv | out-file -Encoding UTF8 -FilePath $filename -Force
	}
	
	
	$FilterButton.Enabled = $true
	$dgv.Enabled = $true
	$StatusLabel.Text = ""
	$ExportButton.Enabled = $true
	$FilterButton.Enabled = $true
	$FindPreviousButton.Enabled = $true
	$ChartButton.Enabled = $true
	$FindNextButton.Enabled = $true
	[System.Windows.Forms.Application]::DoEvents()
	
})
$objForm.Controls.Add($ExportButton)



#Count Label ============================================================
$CountLabel = New-Object System.Windows.Forms.Label
$CountLabel.Location = New-Object System.Drawing.Size(30,90) 
$CountLabel.Size = New-Object System.Drawing.Size(200,15) 
$CountLabel.Text = ""
$CountLabel.TabStop = $false
$CountLabel.ForeColor = [System.Drawing.Color]::Blue
$CountLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$objForm.Controls.Add($CountLabel)

#Data Grid View ============================================================
$dgv = New-Object Windows.Forms.DataGridView
$dgv.Size = New-Object System.Drawing.Size(840,360)
$dgv.Location = New-Object System.Drawing.Size(30,110)
$dgv.AutoGenerateColumns = $false
$dgv.RowHeadersVisible = $false
$dgv.MultiSelect = $false
$dgv.AllowUserToAddRows = $false
$dgv.SelectionMode = [Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dgv.AutoSizeRowsMode = [Windows.Forms.DataGridViewAutoSizeRowsMode]::DisplayedCells  #DisplayedCells AllCells  - DisplayedCells is much better for a large number of rows
$dgv.AutoSizeColumnsMode = [Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill  #DisplayedCells Fill AllCells - Fill is much better for a large number of rows
$dgv.DefaultCellStyle.WrapMode = [Windows.Forms.DataGridViewTriState]::True
$dgv.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom


#ConferenceDateTime, Caller, FeedbackText, Rating, TokenDescription, TokenValue

#$titleColumn0 = New-Object Windows.Forms.DataGridViewImageColumn
$titleColumn0 = New-Object Windows.Forms.DataGridViewTextBoxColumn
$titleColumn0.HeaderText = "Date/Time (GMT)"
$titleColumn0.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
$titleColumn0.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
$titleColumn0.ReadOnly = $true
$titleColumn0.MinimumWidth = 120
$titleColumn0.Width = 120
$dgv.Columns.Add($titleColumn0) | Out-Null


$titleColumn1 = New-Object Windows.Forms.DataGridViewTextBoxColumn
$titleColumn1.HeaderText = "Caller"
$titleColumn1.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
$titleColumn1.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
$titleColumn1.ReadOnly = $true
$titleColumn1.MinimumWidth = 160
$titleColumn1.Width = 160
$dgv.Columns.Add($titleColumn1) | Out-Null


$titleColumn2 = New-Object Windows.Forms.DataGridViewTextBoxColumn
$titleColumn2.HeaderText = "Rating"
$titleColumn2.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
$titleColumn2.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
$titleColumn2.ReadOnly = $true
$titleColumn2.MinimumWidth = 50
$titleColumn2.Width = 50
$dgv.Columns.Add($titleColumn2) | Out-Null


$titleColumn3 = New-Object Windows.Forms.DataGridViewTextBoxColumn
$titleColumn3.HeaderText = "Type"
$titleColumn3.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
$titleColumn3.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
$titleColumn3.ReadOnly = $true
$titleColumn3.MinimumWidth = 50
$titleColumn3.Width = 50
$dgv.Columns.Add($titleColumn3) | Out-Null

$titleColumn4 = New-Object Windows.Forms.DataGridViewTextBoxColumn
$titleColumn4.HeaderText = "Reason(s)"
$titleColumn4.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
$titleColumn4.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
$titleColumn4.ReadOnly = $true
$titleColumn4.MinimumWidth = 150
$titleColumn4.Width = 150
$dgv.Columns.Add($titleColumn4) | Out-Null


$titleColumn5 = New-Object Windows.Forms.DataGridViewTextBoxColumn
$titleColumn5.HeaderText = "FeedbackText"
$titleColumn5.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
$titleColumn5.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
$titleColumn5.ReadOnly = $true
$dgv.Columns.Add($titleColumn5) | Out-Null
$objForm.Controls.Add($dgv)


# $StatusLabel ============================================================
$StatusLabel = New-Object System.Windows.Forms.Label
$StatusLabel.Location = New-Object System.Drawing.Size(20,510) 
$StatusLabel.Size = New-Object System.Drawing.Size(700,15) 
$StatusLabel.Text = ""
$StatusLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$StatusLabel.ForeColor = [System.Drawing.Color]::Red
$StatusLabel.TabStop = $false
$objForm.Controls.Add($StatusLabel)


#$RowNumberLabel ============================================================
$RowNumberLabel = New-Object System.Windows.Forms.Label
$RowNumberLabel.Location = New-Object System.Drawing.Size(790,473) 
$RowNumberLabel.Size = New-Object System.Drawing.Size(200,15) 
$RowNumberLabel.Text = ""
$RowNumberLabel.TabStop = $false
$RowNumberLabel.ForeColor = [System.Drawing.Color]::Green
$RowNumberLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$objForm.Controls.Add($RowNumberLabel)


function ChartDialog([string]$Message, [string]$WindowTitle, [string]$DefaultText)
{
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
    [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")
	
	# Create the Save button.
    $SaveButton = New-Object System.Windows.Forms.Button
    $SaveButton.Location = New-Object System.Drawing.Size(120,450)
    $SaveButton.Size = New-Object System.Drawing.Size(75,25)
    $SaveButton.Text = "Save..."
	$SaveButton.tabIndex = 4
	$SaveButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
    $SaveButton.Add_Click({ 
	
	#File Dialog
	[string] $pathVar = "C:\"
	$Filter="All Files (*.*)|*.*"
	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
	$objDialog = New-Object System.Windows.Forms.SaveFileDialog
	#$objDialog.InitialDirectory = 
	$objDialog.FileName = "Chart.png"
	$objDialog.Filter = $Filter
	$objDialog.Title = "Chart Name"
	$objDialog.CheckFileExists = $false
	$Show = $objDialog.ShowDialog()
	if ($Show -eq "OK")
	{
		$filename = $objDialog.FileName
		Write-Host "Saving: $filename" -foreground "green"
		$Chart.SaveImage($filename, "Png")
	}
	
	})
	
	# Create the Save button.
    $SaveAllButton = New-Object System.Windows.Forms.Button
    $SaveAllButton.Location = New-Object System.Drawing.Size(220,450)
    $SaveAllButton.Size = New-Object System.Drawing.Size(75,25)
    $SaveAllButton.Text = "Save All..."
	$SaveAllButton.tabIndex = 3
	$SaveAllButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
    $SaveAllButton.Add_Click({ 
	
	$objFolderForm = New-Object System.Windows.Forms.FolderBrowserDialog
	$objFolderForm.Description = "Output Image Folder"
	$objFolderForm.SelectedPath = "$ScriptFolder"
	$Show = $objFolderForm.ShowDialog()
	if ($Show -eq "OK")
	{
		$SaveButton.Enabled = $false
		$SaveAllButton.Enabled = $false
		$okButton.Enabled = $false
		
		[string]$filename = $objFolderForm.SelectedPath
		
		#Get the start and end date
		$StartDate = ($StartTimePicker.Value.ToString("yyyy.MM.dd_HH.mm"))
		$EndDate = ($EndTimePicker.Value.ToString("yyyy.MM.dd_HH.mm"))
		
		for($i=0; $i -lt $ChartComboBox.Items.Count; $i++)
		{
			Write-Host "Saving: $filename\Chart-${i}_${StartDate}-${EndDate}.png" -foreground "green"
			$ChartComboBox.SelectedIndex = $i
			$outputfilename = "$filename\Chart-${i}_${StartDate}-${EndDate}.png"
			$Chart.SaveImage($outputfilename, "Png")
		}
		
		$SaveButton.Enabled = $true
		$SaveAllButton.Enabled = $true
		$okButton.Enabled = $true
	}
	
	})
	
	
	# Create the OK button.
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Size(320,450)
    $okButton.Size = New-Object System.Drawing.Size(75,25)
    $okButton.Text = "OK"
	$okButton.tabIndex = 1 
	$okButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
    $okButton.Add_Click({ $form.Close() })

	
	#Chart Label ============================================================
	$ChartLabel = New-Object System.Windows.Forms.Label
	$ChartLabel.Location = New-Object System.Drawing.Size(30,13) 
	$ChartLabel.Size = New-Object System.Drawing.Size(100,15) 
	$ChartLabel.Text = "Select Chart:"
	$ChartLabel.TabStop = $false
	$ChartLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
	$objForm.Controls.Add($ChartLabel)	
	
	#Chart Drop Down box ============================================================
	$ChartComboBox = New-Object System.Windows.Forms.ComboBox 
	$ChartComboBox.Location = New-Object System.Drawing.Size(120,10) 
	$ChartComboBox.Size = New-Object System.Drawing.Size(240,15) 
	$ChartComboBox.DropDownHeight = 150 
	$ChartComboBox.DropDownWidth = 240 
	$ChartComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
	$ChartComboBox.tabIndex = 2
	$ChartComboBox.Sorted = $true
	$ChartComboBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
	$objForm.Controls.Add($ChartComboBox) 

	[void] $ChartComboBox.Items.Add("Stars Bar Graph")
	[void] $ChartComboBox.Items.Add("Stars Pie Graph")
	[void] $ChartComboBox.Items.Add("Reason Pie Graph")
	[void] $ChartComboBox.Items.Add("Reason Bar Graph")
	[void] $ChartComboBox.Items.Add("Type Pie Graph")
	[void] $ChartComboBox.Items.Add("Stars Stacked Bar Graph")
	[void] $ChartComboBox.Items.Add("Trend Over Time Line")
	[void] $ChartComboBox.Items.Add("Top 10 One Star Responders")
	[void] $ChartComboBox.Items.Add("Top 10 Zero Star Users (Lync 2013 Client)")
		
	
		
	# create chart object 
	$Chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart 
	$Chart.Width = 530 
	$Chart.Height = 400 
	$Chart.Left = 20 
	$Chart.Top = 40
	$Chart.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
	
	
	#CHARTING
	$ChartComboBox.Add_SelectedIndexChanged({ 
	
	
		[void]$Chart.Titles.Clear()
		[void]$Chart.ChartAreas.Clear()
		[void]$Chart.Series.Clear()
		[void]$Chart.Legends.Clear()

		#BAR CHART
		if($ChartComboBox.Text -eq "Stars Bar Graph")
		{
			$RowCount = $dgv.Rows.Count
			$theText = ""
			$NoFiveStars = 0
			$NoFourStars = 0
			$NoThreeStars = 0
			$NoTwoStars = 0
			$NoOneStars = 0
			$NoZeroStars = 0
			
			[DateTime]$maxDate = (get-date).addyears(-1000)
			[DateTime]$minDate = (get-date).addyears(1000)
			
			
			for($i=0; $i -lt $RowCount; $i++)
			{
				$theValue = $dgv.Rows[$i].Cells[2].Value
				$theType = $dgv.Rows[$i].Cells[3].Value
				
				
				if($theValue -eq 5)
				{
					$NoFiveStars++
				}
				elseif($theValue -eq 4)
				{
					$NoFourStars++
				}
				elseif($theValue -eq 3)
				{
					$NoThreeStars++
				}
				elseif($theValue -eq 2)
				{
					$NoTwoStars++
				}
				elseif($theValue -eq 1)
				{
					$NoOneStars++
				}
				elseif($theValue -eq 0)
				{
					$NoZeroStars++
				}
				
				#MAX DATE
				[DateTime] $rowDate = Get-Date -Date $dgv.Rows[$i].Cells[0].Value.ToString()
				if($rowDate -gt $maxDate)
				{ 
					$maxDate = $rowDate
				}
				#MIN DATE
				if($rowDate -lt $minDate)
				{ 
					$minDate = $rowDate
				}
			}
			
			#Write-Host "MAX DATE: $maxDate"
			#Write-Host "MIN DATE: $minDate"
			
			# create a chartarea to draw on and add to chart 
			$ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea 
			$Chart.ChartAreas.Add($ChartArea)
			
			# add data to chart 
			$Stars = @{"5 Stars"=$NoFiveStars; "4 Stars"=$NoFourStars; "3 Stars"=$NoThreeStars; "2 Stars"=$NoTwoStars; "1 Stars"=$NoOneStars; "0 Stars"=$NoZeroStars } 
			
			[void]$Chart.Series.Add("Data") 
			$Chart.Series["Data"].Points.DataBindXY($Stars.Keys, $Stars.Values)
			$Chart.Series["Data"].Sort([System.Windows.Forms.DataVisualization.Charting.PointSortOrder]::Ascending, "AxisLabel")
			
			
			$StartDate = $minDate.ToString("yyyy/MM/dd HH:mm") #$StartTimePicker.Value.ToString("yyyy/MM/dd HH:mm")
			$EndDate = $maxDate.ToString("yyyy/MM/dd HH:mm") #$EndTimePicker.Value.ToString("yyyy/MM/dd HH:mm")
			
			if($TimeFormatCheckBox.Checked)
			{
				[void]$Chart.Titles.Add("Call Quality Responses ($StartDate - $EndDate GMT)")
			}
			else
			{
				[void]$Chart.Titles.Add("Call Quality Responses ($StartDate - $EndDate Local Time)")
			}
			 
			
			if($VideoCheckBox.Checked -and $VoiceCheckBox.Checked)
			{
				$ChartArea.AxisX.Title = "No of Stars (Video and Voice)"
			}
			elseif($VoiceCheckBox.Checked)
			{
				$ChartArea.AxisX.Title = "No of Stars (Voice)"
			}
			elseif($VideoCheckBox.Checked)
			{
				$ChartArea.AxisX.Title = "No of Stars (Video)"
			}
			
			$ChartArea.AxisY.Title = "No of Responses"
		}
		
		#PIE CHART
		if($ChartComboBox.Text -eq "Stars Pie Graph")
		{
			$RowCount = $dgv.Rows.Count
			$theText = ""
			$NoFiveStars = 0
			$NoFourStars = 0
			$NoThreeStars = 0
			$NoTwoStars = 0
			$NoOneStars = 0
			$NoZeroStars = 0
			
			[DateTime]$maxDate = (get-date).addyears(-1000)
			[DateTime]$minDate = (get-date).addyears(1000)
			
			
			for($i=0; $i -lt $RowCount; $i++)
			{
				$theValue = $dgv.Rows[$i].Cells[2].Value
				$theType = $dgv.Rows[$i].Cells[3].Value
				
				
				if($theValue -eq 5)
				{
					$NoFiveStars++
				}
				elseif($theValue -eq 4)
				{
					$NoFourStars++
				}
				elseif($theValue -eq 3)
				{
					$NoThreeStars++
				}
				elseif($theValue -eq 2)
				{
					$NoTwoStars++
				}
				elseif($theValue -eq 1)
				{
					$NoOneStars++
				}
				elseif($theValue -eq 0)
				{
					$NoZeroStars++
				}
				
				#MAX DATE
				[DateTime] $rowDate = Get-Date -Date $dgv.Rows[$i].Cells[0].Value.ToString()
				if($rowDate -gt $maxDate)
				{ 
					$maxDate = $rowDate
				}
				#MIN DATE
				if($rowDate -lt $minDate)
				{ 
					$minDate = $rowDate
				}
			}
			

			# create a chartarea to draw on and add to chart 
			$ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea 
			$Chart.ChartAreas.Add($ChartArea)
			
			# add data to chart 
			$Stars = @{"5 Stars"=$NoFiveStars; "4 Stars"=$NoFourStars; "3 Stars"=$NoThreeStars; "2 Stars"=$NoTwoStars; "1 Stars"=$NoOneStars; "0 Stars"=$NoZeroStars } 
			
			#Filter 0's
			$GraphData = @{}
			foreach ($Star in $Stars.Keys) 
			{
				if($Stars[$Star] -gt 0)
				{
					$GraphData.Add($Star, $Stars[$Star])
				}
			}
			
			
			[void]$Chart.Series.Add("Data") 
			$Chart.Series["Data"].Points.DataBindXY($GraphData.Keys, $GraphData.Values)
			# set chart type 
			$Chart.Series["Data"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Pie
			
			$StartDate = $minDate.ToString("yyyy/MM/dd HH:mm") #$StartTimePicker.Value.ToString("yyyy/MM/dd HH:mm")
			$EndDate = $maxDate.ToString("yyyy/MM/dd HH:mm") #$EndTimePicker.Value.ToString("yyyy/MM/dd HH:mm")
			
			#[void]$Chart.Titles.Add("Call Quality Responses ($StartDate - $EndDate)") 
			
			if($TimeFormatCheckBox.Checked)
			{
				[void]$Chart.Titles.Add("Call Quality Responses ($StartDate - $EndDate GMT)")
			}
			else
			{
				[void]$Chart.Titles.Add("Call Quality Responses ($StartDate - $EndDate Local Time)")
			}

			# set chart options 
			$Chart.Series["Data"]["PieLabelStyle"] = "Outside" 
			$Chart.Series["Data"]["PieLineColor"] = "Black" 
			#($Chart.Series["Data"].Points.FindMaxByValue())["Exploded"] = $true
			
			$Chart.Series["Data"].Label = "#VALX (#PERCENT)"
			
		}
		
		#Reason Pie
		if($ChartComboBox.Text -eq "Reason Pie Graph")
		{
			$RowCount = $dgv.Rows.Count
			$theText = ""
			
			#ISSUES LIST INCLUDING C2R CLIENT
			$IssueDescriptions = @{"DistortedSpeech"=0;"ElectronicFeedback"=0;"BackgroundNoise"=0;"MuffledSpeech"=0;"Echo"=0;"FrozenVideo"=0;"PixelatedVideo"=0;"BlurryImage"=0;"PoorColor"=0;"DarkVideo"=0;"NoSpeechNearSide"=0;"NoSpeechFarSide"=0;"IHeardNoise"=0;"VolumeLow"=0;"VoiceCallCutOff"=0;"TalkingOverEachOther"=0;"NoVideoNearSide"=0;"NoVideoFarSide"=0;"PoorQualityVideo"=0;"VideoCutOff"=0;"VideoAudioOutOfSync"=0;}

			
			[DateTime]$maxDate = (get-date).addyears(-1000)
			[DateTime]$minDate = (get-date).addyears(1000)
			for($i=0; $i -lt $RowCount; $i++)
			{
				$theValue = $dgv.Rows[$i].Cells[4].Value
				
				foreach($key in $($IssueDescriptions.keys))
				{
					if($theValue -match $key)
					{
						$IssueDescriptions[$key] = $IssueDescriptions[$key] + 1
					}
				}
				
				
				#MAX DATE
				[DateTime] $rowDate = Get-Date -Date $dgv.Rows[$i].Cells[0].Value.ToString()
				if($rowDate -gt $maxDate)
				{ 
					$maxDate = $rowDate
				}
				#MIN DATE
				if($rowDate -lt $minDate)
				{ 
					$minDate = $rowDate
				}

			}
			
			# create a chartarea to draw on and add to chart 
			$ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea 
			$Chart.ChartAreas.Add($ChartArea)
						
			
			$Reasons = @{}
			foreach ($IssueDescription in $IssueDescriptions.Keys) 
			{
				if($IssueDescriptions[$IssueDescription] -gt 0)
				{
					$Reasons.Add($IssueDescription, $IssueDescriptions[$IssueDescription])
				}
			}

			
			[void]$Chart.Series.Add("Data") 
			$Chart.Series["Data"].Points.DataBindXY($Reasons.Keys, $Reasons.Values)
			#$Chart.Series["Data"].Points.DataBindXY($IssueDescriptions.Keys, $IssueDescriptions.Values)
			# set chart type 
			$Chart.Series["Data"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Pie
			
			$StartDate =  $minDate.ToString("yyyy/MM/dd HH:mm")#$StartTimePicker.Value.ToString("yyyy/MM/dd HH:mm")
			$EndDate = $maxDate.ToString("yyyy/MM/dd HH:mm") #$EndTimePicker.Value.ToString("yyyy/MM/dd HH:mm")
			#[void]$Chart.Titles.Add("Call Quality Responses ($StartDate - $EndDate)") 
			
			if($TimeFormatCheckBox.Checked)
			{
				[void]$Chart.Titles.Add("Call Quality Responses ($StartDate - $EndDate GMT)")
			}
			else
			{
				[void]$Chart.Titles.Add("Call Quality Responses ($StartDate - $EndDate Local Time)")
			}
			

			# set chart options 
			$Chart.Series["Data"]["PieLabelStyle"] = "Outside" 
			$Chart.Series["Data"]["PieLineColor"] = "Black" 
			#($Chart.Series["Data"].Points.FindMaxByValue())["Exploded"] = $true
			
			$Chart.Series["Data"].Label = "#VALX (#PERCENT)" #"#PERCENT{P0}"
			
			<#
			#IF YOU USE THIS VERSION THE GRAPH WILL SHOW A LEGEND WITH ALL VALUES IN IT. THE COLORS ARE TOO MUCH THE SAME THOUGH SO IT'S NOT VERY GOOD.
			if($Reasons.Count -gt 15)
			{
				$legend = New-Object system.Windows.Forms.DataVisualization.Charting.Legend 
				$legend.name = "Legend" 
				$legend.alignment = "Center" 
				$legend.docking = "bottom" 
				#$legend.bordercolor ="orange" 
				$legend.legendstyle = "table" 
				$Chart.Legends.Add($legend)
				
				$chart.Series["Data"].LegendText = "#VALX (#PERCENT)"
				$Chart.Series["Data"].Label = "#PERCENT"
			}
			else
			{
				$Chart.Series["Data"].Label = "#VALX (#PERCENT)" #"#PERCENT{P0}"
			}
			#>
		}
		
		#Reason Bar
		if($ChartComboBox.Text -eq "Reason Bar Graph")
		{
			$RowCount = $dgv.Rows.Count
			$theText = ""
			
			$IssueDescriptions = @{"DistortedSpeech"=0;"ElectronicFeedback"=0;"BackgroundNoise"=0;"MuffledSpeech"=0;"Echo"=0;"FrozenVideo"=0;"PixelatedVideo"=0;"BlurryImage"=0;"PoorColor"=0;"DarkVideo"=0;"NoSpeechNearSide"=0;"NoSpeechFarSide"=0;"IHeardNoise"=0;"VolumeLow"=0;"VoiceCallCutOff"=0;"TalkingOverEachOther"=0;"NoVideoNearSide"=0;"NoVideoFarSide"=0;"PoorQualityVideo"=0;"VideoCutOff"=0;"VideoAudioOutOfSync"=0;}

			
			[DateTime]$maxDate = (get-date).addyears(-1000)
			[DateTime]$minDate = (get-date).addyears(1000)
			for($i=0; $i -lt $RowCount; $i++)
			{
				$theValue = $dgv.Rows[$i].Cells[4].Value
				
				
				foreach($key in $($IssueDescriptions.keys))
				{
					if($theValue -match $key)
					{
						$IssueDescriptions[$key] = $IssueDescriptions[$key] + 1
					}
				}
				
				#MAX DATE
				[DateTime] $rowDate = Get-Date -Date $dgv.Rows[$i].Cells[0].Value.ToString()
				if($rowDate -gt $maxDate)
				{ 
					$maxDate = $rowDate
				}
				#MIN DATE
				if($rowDate -lt $minDate)
				{ 
					$minDate = $rowDate
				}
			}
			
			# create a chartarea to draw on and add to chart 
			$ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea 
			$Chart.ChartAreas.Add($ChartArea)
			$Chart.ChartAreas.AxisX.Interval = 1
			
			# add data to chart 
			#$Reasons = @{"PixelatedVideo"=$PixelatedVideo; "FrozenVideo"=$FrozenVideo; "BlurryImage"=$BlurryImage; "PoorColor"=$PoorColor; "DarkVideo"=$DarkVideo; "BackgroundNoise"=$BackgroundNoise; "MuffledSpeech"=$MuffledSpeech; "Echo"=$Echo; "ElectronicFeedback"=$ElectronicFeedback; "DistortedSpeech"=$DistortedSpeech } 
			

			$Reasons = @{}
			foreach ($IssueDescription in $IssueDescriptions.Keys) 
			{
				if($IssueDescriptions[$IssueDescription] -gt 0)
				{
					$Reasons.Add($IssueDescription, $IssueDescriptions[$IssueDescription])
				}
			}
			
			[void]$Chart.Series.Add("Data") 
			$Chart.Series["Data"].Points.DataBindXY($Reasons.Keys, $Reasons.Values)
			$Chart.Series["Data"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Column
			
			$Chart.Series["Data"].Sort([System.Windows.Forms.DataVisualization.Charting.PointSortOrder]::Ascending, "AxisLabel")
			
			$StartDate = $minDate.ToString("yyyy/MM/dd HH:mm") #$StartTimePicker.Value.ToString("yyyy/MM/dd HH:mm")
			$EndDate = $maxDate.ToString("yyyy/MM/dd HH:mm") #$EndTimePicker.Value.ToString("yyyy/MM/dd HH:mm")
			#[void]$Chart.Titles.Add("Call Quality Responses ($StartDate - $EndDate)") 
			
			if($TimeFormatCheckBox.Checked)
			{
				[void]$Chart.Titles.Add("Call Quality Responses ($StartDate - $EndDate GMT)")
			}
			else
			{
				[void]$Chart.Titles.Add("Call Quality Responses ($StartDate - $EndDate Local Time)")
			}
			
			$ChartArea.AxisY.Title = "No of Responses"
		}
		
		#Type Pie
		if($ChartComboBox.Text -eq "Type Pie Graph")
		{
			$RowCount = $dgv.Rows.Count
			$theText = ""
			$Voice = 0
			$Video = 0

			[DateTime]$maxDate = (get-date).addyears(-1000)
			[DateTime]$minDate = (get-date).addyears(1000)
			
			
			for($i=0; $i -lt $RowCount; $i++)
			{
				$theValue = $dgv.Rows[$i].Cells[3].Value
				
				
				if($theValue -eq "Voice")
				{
					$Voice++
				}
				elseif($theValue -eq "Video")
				{
					$Video++
				}
				
				#MAX DATE
				[DateTime] $rowDate = Get-Date -Date $dgv.Rows[$i].Cells[0].Value.ToString()
				if($rowDate -gt $maxDate)
				{ 
					$maxDate = $rowDate
				}
				#MIN DATE
				if($rowDate -lt $minDate)
				{ 
					$minDate = $rowDate
				}
			}
			

			# create a chartarea to draw on and add to chart 
			$ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea 
			$Chart.ChartAreas.Add($ChartArea)

			$Stars = @{"Voice"=$Voice; "Video"=$Video } 
						
			[void]$Chart.Series.Add("Data") 
			$Chart.Series["Data"].Points.DataBindXY($Stars.Keys, $Stars.Values)
			# set chart type 
			$Chart.Series["Data"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Pie
			
			$StartDate = $minDate.ToString("yyyy/MM/dd HH:mm") #$StartTimePicker.Value.ToString("yyyy/MM/dd HH:mm")
			$EndDate = $maxDate.ToString("yyyy/MM/dd HH:mm") #$EndTimePicker.Value.ToString("yyyy/MM/dd HH:mm")
			#[void]$Chart.Titles.Add("Call Quality Responses ($StartDate - $EndDate)") 
			
			if($TimeFormatCheckBox.Checked)
			{
				[void]$Chart.Titles.Add("Call Quality Responses ($StartDate - $EndDate GMT)")
			}
			else
			{
				[void]$Chart.Titles.Add("Call Quality Responses ($StartDate - $EndDate Local Time)")
			}
			

			# set chart options 
			$Chart.Series["Data"]["PieLabelStyle"] = "Outside" 
			$Chart.Series["Data"]["PieLineColor"] = "Black" 
			#($Chart.Series["Data"].Points.FindMaxByValue())["Exploded"] = $true
			
			$Chart.Series["Data"].Label = "#VALY #VALX (#PERCENT)" #"#PERCENT{P0}"

		}
		
		#Stacked Chart (Comparison Video / Voice) StackedBar
		if($ChartComboBox.Text -eq "Stars Stacked Bar Graph")
		{
			$RowCount = $dgv.Rows.Count
			$theText = ""
			$NoFiveStarsVideo = 0
			$NoFourStarsVideo = 0
			$NoThreeStarsVideo = 0
			$NoTwoStarsVideo = 0
			$NoOneStarsVideo = 0
			$NoZeroStarsVideo = 0
			
			$NoFiveStarsVoice = 0
			$NoFourStarsVoice = 0
			$NoThreeStarsVoice = 0
			$NoTwoStarsVoice = 0
			$NoOneStarsVoice = 0
			$NoZeroStarsVoice = 0
			
			[DateTime]$maxDate = (get-date).addyears(-1000)
			[DateTime]$minDate = (get-date).addyears(1000)
			
			
			for($i=0; $i -lt $RowCount; $i++)
			{
				$theValue = $dgv.Rows[$i].Cells[2].Value
				$theType = $dgv.Rows[$i].Cells[3].Value
				
				
				if($theType -eq "Video")
				{
					if($theValue -eq 5)
					{
						$NoFiveStarsVideo++
					}
					elseif($theValue -eq 4)
					{
						$NoFourStarsVideo++	
					}
					elseif($theValue -eq 3)
					{
						$NoThreeStarsVideo++
					}
					elseif($theValue -eq 2)
					{
						$NoTwoStarsVideo++
					}
					elseif($theValue -eq 1)
					{
						$NoOneStarsVideo++	
					}
					elseif($theValue -eq 0)
					{
						$NoZeroStarsVideo++
					}
				}
				elseif($theType -eq "Voice")
				{
					if($theValue -eq 5)
					{
						$NoFiveStarsVoice++
					}
					elseif($theValue -eq 4)
					{
						$NoFourStarsVoice++
					}
					elseif($theValue -eq 3)
					{
						$NoThreeStarsVoice++
					}
					elseif($theValue -eq 2)
					{
						$NoTwoStarsVoice++
					}
					elseif($theValue -eq 1)
					{
						$NoOneStarsVoice++
					}
					elseif($theValue -eq 0)
					{
						$NoZeroStarsVoice++
					}
				}
				
				#MAX DATE
				[DateTime] $rowDate = Get-Date -Date $dgv.Rows[$i].Cells[0].Value.ToString()
				if($rowDate -gt $maxDate)
				{ 
					$maxDate = $rowDate
				}
				#MIN DATE
				if($rowDate -lt $minDate)
				{ 
					$minDate = $rowDate
				}
			}

			# create a chartarea to draw on and add to chart 
			$ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea 
			$Chart.ChartAreas.Add($ChartArea)
			
			
			# add data to chart 
			$StarsVideo = @{"5 Stars"=$NoFiveStarsVideo; "4 Stars"=$NoFourStarsVideo; "3 Stars"=$NoThreeStarsVideo; "2 Stars"=$NoTwoStarsVideo; "1 Stars"=$NoOneStarsVideo; "0 Stars"=$NoZeroStarsVideo } 
			$StarsVoice = @{"5 Stars"=$NoFiveStarsVoice; "4 Stars"=$NoFourStarsVoice; "3 Stars"=$NoThreeStarsVoice; "2 Stars"=$NoTwoStarsVoice; "1 Stars"=$NoOneStarsVoice; "0 Stars"=$NoZeroStarsVoice } 
			
			
			[void]$Chart.Series.Add("Video Calls") 
			[void]$Chart.Series.Add("Voice Calls") 
			$Chart.Series["Video Calls"].Points.DataBindXY($StarsVideo.Keys, $StarsVideo.Values)
			$Chart.Series["Voice Calls"].Points.DataBindXY($StarsVoice.Keys, $StarsVoice.Values)
			
			$Chart.Series["Video Calls"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::StackedColumn
			$Chart.Series["Voice Calls"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::StackedColumn
						
			$Chart.Series["Video Calls"].Sort([System.Windows.Forms.DataVisualization.Charting.PointSortOrder]::Ascending, "AxisLabel")
			$Chart.Series["Voice Calls"].Sort([System.Windows.Forms.DataVisualization.Charting.PointSortOrder]::Ascending, "AxisLabel")
			
			$StartDate = $minDate.ToString("yyyy/MM/dd HH:mm") #$StartTimePicker.Value.ToString("yyyy/MM/dd HH:mm")
			$EndDate = $maxDate.ToString("yyyy/MM/dd HH:mm") #$EndTimePicker.Value.ToString("yyyy/MM/dd HH:mm")
			#[void]$Chart.Titles.Add("Call Quality Responses ($StartDate - $EndDate)") 
			
			if($TimeFormatCheckBox.Checked)
			{
				[void]$Chart.Titles.Add("Call Quality Responses ($StartDate - $EndDate GMT)")
			}
			else
			{
				[void]$Chart.Titles.Add("Call Quality Responses ($StartDate - $EndDate Local Time)")
			}
			
			
			if($VideoCheckBox.Checked -and $VoiceCheckBox.Checked)
			{
				$ChartArea.AxisX.Title = "No of Stars (Video and Voice)"
			}
			elseif($VoiceCheckBox.Checked)
			{
				$ChartArea.AxisX.Title = "No of Stars (Voice)"
			}
			elseif($VideoCheckBox.Checked)
			{
				$ChartArea.AxisX.Title = "No of Stars (Video)"
			}
			
			$ChartArea.AxisY.Title = "No of Responses"
			
			$Chart.Series["Video Calls"].Label = "#VALY" 
			$Chart.Series["Voice Calls"].Label = "#VALY" 
			
			# legend 
		    $legend = New-Object system.Windows.Forms.DataVisualization.Charting.Legend
		    $legend.name = "Legend1"
		    $Chart.Legends.Add($legend)
			
		}
		
		#Trend Over Time Line
		if($ChartComboBox.Text -eq "Trend Over Time Line")
		{
			$RowCount = $dgv.Rows.Count
			$theText = ""
						
			# add data to chart 
			$Date5Stars = @{} 
			$Date4Stars = @{} 
			$Date3Stars = @{} 
			$Date2Stars = @{} 
			$Date1Stars = @{}
			$Date0Stars = @{}
			
			
			#Intialise all dates in range
			$start = $StartTimePicker.Value
			$finish = $EndTimePicker.Value

			$delta = New-TimeSpan -Days 1

			for ($d = $start; $d -le $finish; $d +=$delta) 
			{
				[string]$aDay = $d.toString("yyyy/MM/dd")
				
				$Date0Stars.Add($aDay, 0)
				$Date1Stars.Add($aDay, 0)
				$Date2Stars.Add($aDay, 0)
				$Date3Stars.Add($aDay, 0)
				$Date4Stars.Add($aDay, 0)
				$Date5Stars.Add($aDay, 0)
			}
			
			[DateTime]$maxDate = (get-date).addyears(-1000)
			[DateTime]$minDate = (get-date).addyears(1000)
			
						
			for($i=0; $i -lt $RowCount; $i++)
			{
				$theValue = $dgv.Rows[$i].Cells[2].Value
				$theType = $dgv.Rows[$i].Cells[3].Value
				[string]$theDate = $dgv.Rows[$i].Cells[0].Value
				
				$theDateArray = $theDate.Split(" ")
				[string]$theDay = $theDateArray[0]
				
				if($theValue -eq 0)
				{
					if($Date0Stars.ContainsKey($theDay))
					{
						$Date0Stars[$theDay] = $Date0Stars[$theDay] + 1
					}
				}
				elseif($theValue -eq 1)
				{
					if($Date1Stars.ContainsKey($theDay))
					{
						$Date1Stars[$theDay] = $Date1Stars[$theDay] + 1
					}
				}
				elseif($theValue -eq 2)
				{
					if($Date2Stars.ContainsKey($theDay))
					{
						$Date2Stars[$theDay] = $Date2Stars[$theDay] + 1
					}
				}
				elseif($theValue -eq 3)
				{
					if($Date3Stars.ContainsKey($theDay))
					{
						$Date3Stars[$theDay] = $Date3Stars[$theDay] + 1
					}
				}
				elseif($theValue -eq 4)
				{
					if($Date4Stars.ContainsKey($theDay))
					{
						$Date4Stars[$theDay] = $Date4Stars[$theDay] + 1
					}
				}
				elseif($theValue -eq 5)
				{
					if($Date5Stars.ContainsKey($theDay))
					{
						$Date5Stars[$theDay] = $Date5Stars[$theDay] + 1
					}
				}
				
				#MAX DATE
				[DateTime] $rowDate = Get-Date -Date $dgv.Rows[$i].Cells[0].Value.ToString()
				if($rowDate -gt $maxDate)
				{ 
					$maxDate = $rowDate
				}
				#MIN DATE
				if($rowDate -lt $minDate)
				{ 
					$minDate = $rowDate
				}
				
			}

			# create a chartarea to draw on and add to chart 
			$ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea 
			$Chart.ChartAreas.Add($ChartArea)
			
			
			[void]$Chart.Series.Add("0 Stars")
			[void]$Chart.Series.Add("1 Stars") 
			[void]$Chart.Series.Add("2 Stars") 
			[void]$Chart.Series.Add("3 Stars") 
			[void]$Chart.Series.Add("4 Stars") 
			[void]$Chart.Series.Add("5 Stars") 
			$Chart.Series["0 Stars"].Points.DataBindXY($Date0Stars.Keys, $Date0Stars.Values)
			$Chart.Series["1 Stars"].Points.DataBindXY($Date1Stars.Keys, $Date1Stars.Values)
			$Chart.Series["2 Stars"].Points.DataBindXY($Date2Stars.Keys, $Date2Stars.Values)
			$Chart.Series["3 Stars"].Points.DataBindXY($Date3Stars.Keys, $Date3Stars.Values)
			$Chart.Series["4 Stars"].Points.DataBindXY($Date4Stars.Keys, $Date4Stars.Values)
			$Chart.Series["5 Stars"].Points.DataBindXY($Date5Stars.Keys, $Date5Stars.Values)
			
			$Chart.Series["0 Stars"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
			$Chart.Series["1 Stars"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
			$Chart.Series["2 Stars"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
			$Chart.Series["3 Stars"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
			$Chart.Series["4 Stars"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
			$Chart.Series["5 Stars"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
			
			$Chart.Series["0 Stars"].MarkerStyle = [System.Windows.Forms.DataVisualization.Charting.MarkerStyle]::Circle
			$Chart.Series["1 Stars"].MarkerStyle = [System.Windows.Forms.DataVisualization.Charting.MarkerStyle]::Cross
			$Chart.Series["2 Stars"].MarkerStyle = [System.Windows.Forms.DataVisualization.Charting.MarkerStyle]::Diamond
			$Chart.Series["3 Stars"].MarkerStyle = [System.Windows.Forms.DataVisualization.Charting.MarkerStyle]::Square
			$Chart.Series["4 Stars"].MarkerStyle = [System.Windows.Forms.DataVisualization.Charting.MarkerStyle]::Triangle
			$Chart.Series["5 Stars"].MarkerStyle = [System.Windows.Forms.DataVisualization.Charting.MarkerStyle]::Star4
			
			$Chart.Series["0 Stars"].MarkerSize = 7
			$Chart.Series["1 Stars"].MarkerSize = 7
			$Chart.Series["2 Stars"].MarkerSize = 7
			$Chart.Series["3 Stars"].MarkerSize = 7
			$Chart.Series["4 Stars"].MarkerSize = 7
			$Chart.Series["5 Stars"].MarkerSize = 7
			
			
			$Chart.ChartAreas[0].CursorX.IsUserSelectionEnabled = $true
			$Chart.ChartAreas[0].AxisX.ScaleView.Zoomable = $true
			$Chart.ChartAreas[0].AxisX.ScrollBar.IsPositionedInside = $false
			$Chart.ChartAreas[0].AxisX.ScrollBar.BackColor = [System.Drawing.Color]::White
			$Chart.ChartAreas[0].AxisX.ScrollBar.LineColor = [System.Drawing.Color]::Black
			$Chart.ChartAreas[0].AxisX.ScrollBar.ButtonColor = [System.Drawing.Color]::White
			
			$Chart.ChartAreas[0].CursorY.IsUserSelectionEnabled = $true
			$Chart.ChartAreas[0].AxisY.ScaleView.Zoomable = $true
			$Chart.ChartAreas[0].AxisY.ScrollBar.IsPositionedInside = $false
			$Chart.ChartAreas[0].AxisY.ScrollBar.BackColor = [System.Drawing.Color]::White
			$Chart.ChartAreas[0].AxisY.ScrollBar.LineColor = [System.Drawing.Color]::Black
			$Chart.ChartAreas[0].AxisY.ScrollBar.ButtonColor = [System.Drawing.Color]::White
			
			
			$Chart.Series["0 Stars"].Sort([System.Windows.Forms.DataVisualization.Charting.PointSortOrder]::Ascending, "AxisLabel")
			$Chart.Series["1 Stars"].Sort([System.Windows.Forms.DataVisualization.Charting.PointSortOrder]::Ascending, "AxisLabel")
			$Chart.Series["2 Stars"].Sort([System.Windows.Forms.DataVisualization.Charting.PointSortOrder]::Ascending, "AxisLabel")
			$Chart.Series["3 Stars"].Sort([System.Windows.Forms.DataVisualization.Charting.PointSortOrder]::Ascending, "AxisLabel")
			$Chart.Series["4 Stars"].Sort([System.Windows.Forms.DataVisualization.Charting.PointSortOrder]::Ascending, "AxisLabel")
			$Chart.Series["5 Stars"].Sort([System.Windows.Forms.DataVisualization.Charting.PointSortOrder]::Ascending, "AxisLabel")
			
			$StartDate = $minDate.ToString("yyyy/MM/dd HH:mm") #$StartTimePicker.Value.ToString("yyyy/MM/dd HH:mm")
			$EndDate = $maxDate.ToString("yyyy/MM/dd HH:mm") #$EndTimePicker.Value.ToString("yyyy/MM/dd HH:mm")
			#[void]$Chart.Titles.Add("Call Quality Responses ($StartDate - $EndDate)") 
			
			if($TimeFormatCheckBox.Checked)
			{
				[void]$Chart.Titles.Add("Call Quality Responses ($StartDate - $EndDate GMT)")
			}
			else
			{
				[void]$Chart.Titles.Add("Call Quality Responses ($StartDate - $EndDate Local Time)")
			}
			
			
			$ChartArea.AxisY.Title = "No of Stars"
			
			# legend 
		    $legend = New-Object system.Windows.Forms.DataVisualization.Charting.Legend
		    $legend.name = "Legend1"
		    $Chart.Legends.Add($legend)
			
		}
		
		#Unhappy Users Graph
		if($ChartComboBox.Text -eq "Top 10 One Star Responders")
		{
			$RowCount = $dgv.Rows.Count
			$theText = ""
			
			[DateTime]$maxDate = (get-date).addyears(-1000)
			[DateTime]$minDate = (get-date).addyears(1000)
						
			$UserHash = @{}
			
			for($i=0; $i -lt $RowCount; $i++)
			{
				[string]$caller = $dgv.Rows[$i].Cells[1].Value
				
				$caller = $caller.Replace("sip:", "")
				
				if($UserHash[$caller] -eq $null)
				{
					$UserHash.add($caller, 0)
				}
			}
			
			for($i=0; $i -lt $RowCount; $i++)
			{
				$caller = $dgv.Rows[$i].Cells[1].Value
				$theValue = $dgv.Rows[$i].Cells[2].Value
								
				$caller = $caller.Replace("sip:", "")				
				$CurrentCount = $UserHash[$caller]
				
				[int]$CurrentCountInt = 0
				[bool]$result = [int]::TryParse($CurrentCount, [ref]$CurrentCountInt)
				
				if($result)
				{
					if($theValue -eq 1)
					{
						$CurrentCountInt++
						$UserHash[$caller] = $CurrentCountInt
					}
				}
				
			
				#MAX DATE
				[DateTime] $rowDate = Get-Date -Date $dgv.Rows[$i].Cells[0].Value.ToString()
				if($rowDate -gt $maxDate)
				{ 
					$maxDate = $rowDate
				}
				#MIN DATE
				if($rowDate -lt $minDate)
				{ 
					$minDate = $rowDate
				}
			}
			
			Write-Host "All Users Number of One Star Ratings" -ForegroundColor "Yellow"
			$str = $UserHash | Out-String
			Write-Host $str -ForegroundColor "Yellow"
						
			$UserHashTemp = $UserHash.GetEnumerator() | Where-Object {$_.Value -gt 0} | Sort-Object -Property Value -Descending | Select-Object -First 10
			$UserHashTruncate = @{}
			foreach ($p in $UserHashTemp)
			{
				$UserHashTruncate.add($p.name,$p.value)
			}
			
			Write-Host "Top 10 One Star Users Ranked" -ForegroundColor "Yellow"
			$str = $UserHashTemp | Out-String
			Write-Host $str -ForegroundColor "Yellow"

			
			#Write-Host "MAX DATE: $maxDate"
			#Write-Host "MIN DATE: $minDate"
			
			# create a chartarea to draw on and add to chart 
			$ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
			$Chart.ChartAreas.Add($ChartArea)
			$Chart.ChartAreas.AxisX.Interval = 1
			
			[void]$Chart.Series.Add("Data") 
			$Chart.Series["Data"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Column
			$Chart.Series["Data"].Points.DataBindXY($UserHashTruncate.Keys, $UserHashTruncate.Values)
			$Chart.Series["Data"].Sort([System.Windows.Forms.DataVisualization.Charting.PointSortOrder]::Descending, "Y") #"AxisLabel"
			$Chart.Series["Data"].Color = "Red"
			
			$StartDate = $minDate.ToString("yyyy/MM/dd HH:mm") #$StartTimePicker.Value.ToString("yyyy/MM/dd HH:mm")
			$EndDate = $maxDate.ToString("yyyy/MM/dd HH:mm") #$EndTimePicker.Value.ToString("yyyy/MM/dd HH:mm")
			
			if($TimeFormatCheckBox.Checked)
			{
				[void]$Chart.Titles.Add("Top 10 One Star Responders ($StartDate - $EndDate GMT)")
			}
			else
			{
				[void]$Chart.Titles.Add("Top 10 One Star Responders ($StartDate - $EndDate Local Time)")
			}
			 			
			$ChartArea.AxisX.Title = "User"
			$ChartArea.AxisY.Title = "No of One Star Responses"
		}
		
		#Users that are not responding (and that use the Lync 2013 client with SfB UI)
		if($ChartComboBox.Text -eq "Top 10 Zero Star Users (Lync 2013 Client)")
		{
			$RowCount = $dgv.Rows.Count
			$theText = ""
			
			[DateTime]$maxDate = (get-date).addyears(-1000)
			[DateTime]$minDate = (get-date).addyears(1000)
						
			$UserHash = @{}
			
			for($i=0; $i -lt $RowCount; $i++)
			{
				[string]$caller = $dgv.Rows[$i].Cells[1].Value
				
				$caller = $caller.Replace("sip:", "")
				
				if($UserHash[$caller] -eq $null)
				{
					$UserHash.add($caller, 0)
				}
			}
			
			for($i=0; $i -lt $RowCount; $i++)
			{
				$caller = $dgv.Rows[$i].Cells[1].Value
				$theValue = $dgv.Rows[$i].Cells[2].Value
								
				$caller = $caller.Replace("sip:", "")				
				$CurrentCount = $UserHash[$caller]
				
				[int]$CurrentCountInt = 0
				[bool]$result = [int]::TryParse($CurrentCount, [ref]$CurrentCountInt)
				
				if($result)
				{
					if($theValue -eq 0)
					{
						$CurrentCountInt++
						$UserHash[$caller] = $CurrentCountInt
					}
				}
				
			
				#MAX DATE
				[DateTime] $rowDate = Get-Date -Date $dgv.Rows[$i].Cells[0].Value.ToString()
				if($rowDate -gt $maxDate)
				{ 
					$maxDate = $rowDate
				}
				#MIN DATE
				if($rowDate -lt $minDate)
				{ 
					$minDate = $rowDate
				}
			}
			
			Write-Host "All Users Number of Zero Star Ratings" -ForegroundColor "Yellow"
			$str = $UserHash | Out-String
			Write-Host $str -ForegroundColor "Yellow"
						
			$UserHashTemp = $UserHash.GetEnumerator() | Where-Object {$_.Value -gt 0} | Sort-Object -Property Value -Descending | Select-Object -First 10
			$UserHashTruncate = @{}
			foreach ($p in $UserHashTemp)
			{
				$UserHashTruncate.add($p.name,$p.value)
			}
			
			Write-Host "Top 10 Zero Star Users Ranked" -ForegroundColor "Yellow"
			$str = $UserHashTemp | Out-String
			Write-Host $str -ForegroundColor "Yellow"

			
			#Write-Host "MAX DATE: $maxDate"
			#Write-Host "MIN DATE: $minDate"
			
			# create a chartarea to draw on and add to chart 
			$ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
			$Chart.ChartAreas.Add($ChartArea)
			$Chart.ChartAreas.AxisX.Interval = 1
			
			[void]$Chart.Series.Add("Data") 
			$Chart.Series["Data"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Column
			$Chart.Series["Data"].Points.DataBindXY($UserHashTruncate.Keys, $UserHashTruncate.Values)
			$Chart.Series["Data"].Sort([System.Windows.Forms.DataVisualization.Charting.PointSortOrder]::Descending, "Y") #"AxisLabel"
			$Chart.Series["Data"].Color = "Red"
			
			$StartDate = $minDate.ToString("yyyy/MM/dd HH:mm") #$StartTimePicker.Value.ToString("yyyy/MM/dd HH:mm")
			$EndDate = $maxDate.ToString("yyyy/MM/dd HH:mm") #$EndTimePicker.Value.ToString("yyyy/MM/dd HH:mm")
			
			if($TimeFormatCheckBox.Checked)
			{
				[void]$Chart.Titles.Add("Top 10 Zero Star Responders [Lync 2013 Client] ($StartDate - $EndDate GMT)")
			}
			else
			{
				[void]$Chart.Titles.Add("Top 10 Zero Star Responders [Lync 2013 Client] ($StartDate - $EndDate Local Time)")
			}
			 			
			$ChartArea.AxisX.Title = "User"
			$ChartArea.AxisY.Title = "No of Zero Star Responses"
		}
				
		
	})

	$ChartComboBox.SelectedIndex = $ChartComboBox.FindStringExact("Stars Bar Graph")
	
	 
    # Create the form.
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = $WindowTitle
    $form.Size = New-Object System.Drawing.Size(580,530)
    $form.MinimumSize = New-Object System.Drawing.Size(580,530) 
	$form.StartPosition = "CenterScreen"
	[byte[]]$WindowIcon = @(137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 32, 0, 0, 0, 32, 8, 6, 0, 0, 0, 115, 122, 122, 244, 0, 0, 0, 6, 98, 75, 71, 68, 0, 255, 0, 255, 0, 255, 160, 189, 167, 147, 0, 0, 0, 9, 112, 72, 89, 115, 0, 0, 11, 19, 0, 0, 11, 19, 1, 0, 154, 156, 24, 0, 0, 0, 7, 116, 73, 77, 69, 7, 225, 7, 26, 1, 36, 51, 211, 178, 227, 235, 0, 0, 5, 235, 73, 68, 65, 84, 88, 195, 197, 151, 91, 108, 92, 213, 21, 134, 191, 189, 207, 57, 115, 159, 216, 78, 176, 27, 98, 72, 226, 88, 110, 66, 66, 34, 185, 161, 168, 193, 73, 21, 17, 2, 2, 139, 75, 164, 182, 106, 145, 170, 190, 84, 74, 104, 65, 16, 144, 218, 138, 138, 244, 173, 69, 106, 101, 42, 129, 42, 149, 170, 162, 15, 168, 168, 151, 7, 4, 22, 180, 1, 41, 92, 172, 52, 196, 68, 105, 130, 19, 138, 98, 76, 154, 27, 174, 227, 248, 58, 247, 57, 103, 175, 62, 236, 241, 177, 199, 246, 140, 67, 26, 169, 251, 237, 236, 61, 179, 215, 191, 214, 191, 214, 191, 214, 86, 188, 62, 37, 252, 31, 151, 174, 123, 42, 224, 42, 72, 56, 138, 152, 99, 191, 175, 247, 114, 107, 29, 172, 75, 106, 94, 254, 74, 156, 109, 13, 58, 180, 155, 53, 240, 216, 64, 129, 63, 156, 43, 95, 55, 0, 106, 62, 5, 158, 134, 83, 59, 147, 116, 36, 106, 7, 103, 188, 44, 228, 13, 120, 202, 126, 151, 12, 100, 3, 225, 183, 231, 203, 60, 55, 88, 66, 4, 80, 215, 0, 96, 89, 68, 113, 97, 87, 138, 180, 3, 163, 101, 120, 116, 160, 192, 161, 81, 159, 203, 69, 33, 230, 40, 58, 27, 52, 251, 215, 69, 248, 198, 74, 183, 238, 165, 175, 141, 248, 60, 114, 178, 192, 165, 188, 44, 9, 100, 22, 128, 192, 127, 238, 73, 209, 18, 81, 252, 109, 52, 224, 222, 247, 179, 179, 46, 206, 93, 102, 142, 119, 193, 76, 216, 96, 247, 13, 46, 223, 189, 201, 101, 207, 74, 143, 148, 99, 183, 159, 250, 184, 72, 207, 96, 169, 46, 136, 16, 192, 183, 91, 61, 94, 233, 140, 241, 81, 198, 176, 229, 173, 204, 226, 198, 175, 102, 5, 194, 243, 157, 113, 246, 221, 236, 225, 42, 232, 29, 9, 184, 255, 104, 174, 62, 0, 165, 192, 239, 78, 163, 129, 174, 195, 57, 14, 143, 5, 255, 115, 114, 197, 29, 197, 200, 221, 41, 82, 14, 188, 63, 30, 240, 245, 190, 220, 162, 145, 208, 0, 141, 174, 66, 1, 37, 129, 195, 163, 254, 34, 40, 1, 191, 70, 25, 250, 50, 75, 197, 156, 149, 15, 132, 27, 254, 62, 205, 229, 178, 176, 163, 201, 161, 103, 115, 172, 182, 14, 196, 181, 53, 114, 38, 107, 64, 22, 194, 92, 147, 80, 200, 67, 105, 50, 247, 165, 171, 156, 104, 141, 105, 70, 186, 211, 200, 131, 105, 214, 46, 82, 53, 69, 3, 119, 244, 217, 240, 63, 177, 214, 35, 233, 170, 250, 66, 164, 20, 11, 221, 52, 240, 171, 77, 49, 114, 6, 198, 74, 18, 158, 106, 5, 239, 110, 79, 208, 236, 41, 254, 93, 16, 206, 102, 204, 162, 30, 14, 78, 27, 158, 60, 93, 68, 1, 7, 191, 150, 176, 73, 60, 31, 64, 182, 178, 185, 49, 169, 103, 80, 132, 235, 166, 164, 38, 238, 64, 66, 67, 104, 94, 224, 229, 206, 56, 111, 93, 182, 116, 61, 246, 81, 177, 118, 166, 107, 248, 253, 121, 43, 92, 119, 52, 106, 86, 39, 245, 66, 0, 147, 101, 9, 105, 188, 171, 165, 186, 198, 127, 179, 57, 202, 233, 233, 106, 216, 9, 79, 113, 169, 96, 216, 119, 179, 135, 47, 112, 240, 114, 185, 110, 169, 77, 149, 132, 95, 159, 181, 32, 182, 54, 58, 139, 83, 112, 231, 7, 121, 0, 126, 210, 17, 129, 96, 150, 134, 213, 9, 205, 84, 185, 42, 29, 121, 103, 91, 130, 15, 38, 45, 228, 105, 95, 40, 207, 97, 173, 209, 83, 124, 179, 213, 227, 153, 13, 81, 16, 91, 205, 247, 174, 116, 113, 42, 118, 31, 89, 227, 86, 37, 109, 8, 224, 189, 97, 159, 178, 64, 71, 82, 207, 166, 129, 192, 75, 231, 203, 180, 68, 170, 235, 252, 95, 57, 195, 150, 138, 218, 156, 43, 8, 70, 102, 43, 98, 96, 103, 146, 63, 119, 198, 120, 115, 216, 210, 243, 179, 245, 81, 222, 248, 106, 156, 141, 73, 77, 201, 192, 109, 141, 14, 86, 171, 231, 39, 161, 99, 209, 158, 43, 152, 48, 156, 237, 41, 205, 123, 163, 1, 174, 99, 55, 38, 3, 225, 209, 142, 40, 7, 78, 23, 217, 182, 220, 2, 120, 247, 202, 172, 59, 27, 155, 28, 90, 163, 138, 76, 32, 28, 159, 12, 192, 23, 30, 110, 181, 148, 238, 63, 85, 64, 128, 166, 121, 149, 160, 23, 118, 96, 21, 122, 255, 226, 150, 40, 103, 178, 134, 132, 182, 123, 167, 50, 134, 95, 222, 18, 229, 108, 198, 112, 99, 212, 238, 29, 155, 156, 5, 240, 253, 53, 54, 84, 127, 25, 246, 9, 4, 214, 175, 112, 104, 139, 107, 46, 20, 132, 129, 41, 179, 196, 60, 96, 108, 228, 155, 61, 107, 60, 237, 41, 140, 82, 100, 138, 66, 186, 146, 151, 67, 89, 195, 119, 142, 231, 65, 36, 212, 251, 209, 188, 132, 212, 116, 85, 18, 236, 233, 143, 139, 0, 252, 174, 34, 62, 71, 39, 131, 80, 107, 138, 82, 11, 128, 182, 213, 176, 33, 169, 33, 128, 159, 174, 143, 176, 231, 104, 30, 20, 172, 170, 120, 187, 111, 181, 199, 171, 151, 124, 80, 48, 94, 17, 204, 111, 173, 246, 160, 44, 188, 182, 45, 73, 103, 131, 189, 110, 120, 218, 240, 192, 74, 151, 29, 77, 22, 80, 207, 80, 137, 6, 79, 227, 42, 136, 42, 112, 230, 244, 153, 16, 128, 18, 155, 193, 0, 127, 237, 74, 48, 81, 18, 50, 190, 128, 8, 55, 198, 236, 207, 186, 251, 243, 161, 10, 205, 112, 255, 189, 85, 46, 178, 103, 25, 61, 67, 37, 222, 24, 177, 168, 142, 237, 74, 209, 28, 213, 76, 248, 66, 206, 192, 67, 95, 242, 56, 240, 229, 8, 253, 21, 26, 126, 176, 54, 178, 112, 34, 18, 5, 63, 255, 180, 196, 211, 237, 17, 20, 240, 236, 39, 37, 11, 79, 89, 158, 247, 159, 242, 57, 50, 211, 164, 20, 60, 126, 178, 64, 68, 131, 163, 96, 239, 201, 2, 34, 112, 100, 220, 231, 135, 107, 35, 188, 114, 209, 103, 119, 179, 67, 163, 171, 24, 200, 24, 122, 134, 138, 124, 158, 23, 86, 197, 53, 23, 239, 74, 242, 112, 171, 199, 243, 131, 69, 112, 212, 188, 137, 40, 0, 121, 48, 109, 109, 244, 102, 174, 105, 8, 92, 151, 208, 244, 109, 79, 112, 177, 32, 220, 182, 76, 115, 123, 95, 142, 254, 137, 32, 188, 127, 172, 59, 133, 163, 160, 225, 245, 105, 112, 213, 188, 42, 112, 224, 197, 138, 108, 158, 216, 153, 248, 226, 61, 88, 224, 79, 91, 227, 180, 189, 157, 97, 115, 74, 115, 104, 44, 160, 127, 78, 153, 162, 160, 28, 64, 84, 171, 218, 101, 184, 247, 159, 5, 174, 248, 176, 37, 165, 121, 118, 83, 244, 11, 5, 161, 179, 209, 225, 76, 222, 240, 194, 230, 24, 142, 134, 61, 253, 121, 112, 170, 69, 172, 33, 162, 24, 47, 75, 157, 177, 92, 65, 87, 95, 22, 128, 31, 183, 69, 56, 176, 33, 90, 37, 205, 245, 214, 241, 241, 128, 67, 35, 1, 39, 38, 13, 94, 239, 52, 147, 229, 234, 255, 221, 211, 234, 17, 85, 208, 119, 37, 176, 237, 116, 177, 169, 120, 38, 148, 91, 151, 59, 124, 216, 149, 168, 12, 153, 1, 123, 79, 228, 25, 206, 203, 82, 47, 137, 186, 244, 100, 187, 211, 36, 52, 220, 255, 97, 158, 222, 138, 84, 235, 26, 131, 26, 199, 198, 3, 154, 14, 102, 152, 240, 133, 7, 90, 28, 62, 223, 157, 226, 165, 173, 113, 86, 120, 138, 168, 14, 29, 176, 169, 163, 150, 54, 254, 199, 219, 227, 36, 52, 156, 206, 25, 122, 47, 148, 107, 191, 11, 22, 72, 165, 130, 95, 108, 140, 241, 163, 54, 111, 230, 46, 138, 6, 2, 17, 130, 202, 212, 173, 21, 228, 12, 220, 249, 143, 28, 3, 19, 166, 170, 53, 183, 196, 20, 71, 182, 39, 105, 139, 219, 205, 230, 131, 25, 70, 75, 114, 245, 0, 102, 100, 122, 69, 76, 177, 171, 217, 229, 153, 142, 8, 183, 166, 106, 243, 112, 46, 47, 97, 146, 165, 92, 104, 175, 140, 106, 99, 62, 108, 122, 39, 195, 112, 65, 234, 191, 140, 150, 10, 37, 70, 64, 43, 54, 164, 53, 77, 17, 133, 8, 92, 42, 26, 118, 44, 119, 121, 170, 61, 66, 103, 186, 26, 220, 80, 78, 120, 238, 179, 18, 47, 12, 150, 170, 43, 226, 154, 0, 92, 197, 155, 0, 20, 237, 203, 172, 238, 127, 50, 101, 108, 239, 175, 147, 36, 238, 117, 125, 234, 86, 12, 125, 58, 51, 100, 106, 150, 124, 36, 254, 23, 153, 41, 93, 205, 81, 212, 105, 60, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130)
	$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
	$form.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
	$form.Topmost = $True
    $form.AcceptButton = $okButton
    $form.ShowInTaskbar = $true
     
		
	$form.Controls.Add($Chart)
	$form.Controls.Add($SaveButton)
	$form.Controls.Add($SaveAllButton)
	$form.Controls.Add($okButton)
	$form.Controls.Add($ChartComboBox)
	$form.Controls.Add($ChartLabel)
	
	
    # Initialize and show the form.
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() > $null   # Trash the text of the button that was clicked.
     
    # Return the text that the user entered.
    return $form.Tag
}


function Get-CallQualityEvents
{
	
	$DatabaseServers = $Null
	
	$ServerName = $MonitoringServerDropDown.SelectedItem
	
	if($ServerName -eq "ALL")
	{
		$DatabaseServers = Get-CSService -MonitoringDatabase  | Select-Object Identity,SqlInstanceName,Version
	}
	else
	{
		$DatabaseServers = Get-CSService -Identity "MonitoringDatabase:${ServerName}"
	}
	
	[DateTime]$StartHour = $StartTimePicker.Value
	[DateTime]$EndHour = $EndTimePicker.Value
	
	$dgv.Rows.Clear()
	
	if($EndHour -gt $StartHour)
	{
		
		if($DatabaseServers -eq $null)
		{
			Write-Host "No Monitoring Database found in this Skype for Business environment..." -foreground "red"
		}
		else
		{
			foreach($DatabaseServer in $DatabaseServers)  #Check all of the databases
			{
				$sqlconnecterror = $false
				[string]$Server = $DatabaseServer.Identity
				[string]$SQLInstance = $DatabaseServer.SqlInstanceName
				[string]$SQLVersion = $DatabaseServer.Version   # Must be at least 7 for Skype for Business
				
				$Server = $Server.Replace("MonitoringDatabase:","")
				
				$SQLVersion
				
				[int]$intNum = [convert]::ToInt32($SQLVersion, 10)
				Write-Host "INFO: SQL Version: $intNum" -foreground "yellow"
				if($intNum -ge 7)
				{
					Write-Host "Connecting to Monitoring server: $Server Instance: $SQLInstance" -foreground "Yellow"
					
					#Define SQL Connection String
					[string]$connstring = "server=$Server\$SQLInstance;database=QoEMetrics;trusted_connection=true;"
				 
					#Define SQL Command
					[object]$command = New-Object System.Data.SqlClient.SqlCommand

					
					#NEW - LOOKING UP ALL
					Write-Host "QUERYING FOR ALL RECORDS"
					
					[DateTime]$StartHour = $StartTimePicker.Value
					[DateTime]$EndHour = $EndTimePicker.Value
					
					if(!$TimeFormatCheckBox.Checked)
					{
						[DateTime]$StartHour = $StartHour.ToUniversalTime()
						[DateTime]$EndHour = $EndHour.ToUniversalTime()
					}
					
					#SQL YYYY-MM-DD HH:MI:SS
					[string]$StartDate = $StartHour.ToString("yyyy-MM-dd HH:mm:ss.fff") # HH:mm:ss
					[string]$EndDate = $EndHour.ToString("yyyy-MM-dd HH:mm:ss.fff")
					
					Write-Host "SQL Start Date: $StartDate"
					Write-Host "SQL End Date: $EndDate"
				
				
				#SQL Query
				$command.CommandText = "SELECT
	s.ConferenceDateTime
	,Caller.URI as Caller
	,CallerCqf.FeedbackText 
	,CallerCqf.Rating
	,CallerCqfToken.TokenValue
	,CallerCqfToken.TokenId
FROM [Session] s WITH (NOLOCK)
	INNER JOIN [MediaLine] AS m WITH (NOLOCK) ON 
		m.ConferenceDateTime = s.ConferenceDateTime
		AND m.SessionSeq = s.SessionSeq                        
	INNER JOIN [CallQualityFeedback] AS CallerCqf WITH (NOLOCK) ON
		CallerCqf.ConferenceDateTime  = s.ConferenceDateTime 
		and
		CallerCqf.SessionSeq = s.SessionSeq 
	INNER JOIN [CallQualityFeedbackToken] AS CallerCqfToken WITH (NOLOCK) ON
		CallerCqfToken.ConferenceDateTime  = s.ConferenceDateTime 
		and
		CallerCqfToken.SessionSeq = s.SessionSeq
		and
		CallerCqfToken.FromURI = CallerCqf.FromURI
	INNER JOIN [User] AS Caller WITH (NOLOCK) ON
		Caller.UserKey = CallerCqf.FromURI
	WHERE s.ConferenceDateTime BETWEEN {ts `'$StartDate`'} AND {ts `'$EndDate`'}"
						
					[object]$connection = New-Object System.Data.SqlClient.SqlConnection
					$connection.ConnectionString = $connstring
					try {
					$connection.Open()
					} catch [Exception] {
						write-host ""
						write-host "WARNING: Skype4B Call Quality Viewer was unable to connect to database $server\$SQLInstance. Note: This error is expected if this is a secondary SQL mirrored database. If this is a primary database, please check that the server is online. Also check that UDP 1434 and the Dynamic SQL TCP Port for the Lync/Skype4B Named Instance are open in the Windows Firewall on $server." -foreground "red"
						write-host ""
						#$StatusLabel.Text = "Error connecting to $server. Refer to Powershell window."
						$sqlconnecterror = $true
					}
					
					$tempstoreVoice = @()
					if(!$sqlconnecterror)
					{
						$command.Connection = $connection
										 
						[object]$sqladapter = New-Object System.Data.SqlClient.SqlDataAdapter
						$sqladapter.SelectCommand = $command
					 
						[object]$results = New-Object System.Data.Dataset
						try {
						$recordcount = $sqladapter.Fill($results)
						} catch [Exception] {
							write-host ""
							write-host "Error running SQL on $server : $_" -foreground "red"
							write-host ""
						}
						
						$tempstoreVoice = $results.Tables[0].rows
					}
					$connection.Close()
					
												
					
					
					[DateTime]$StartHour = $StartTimePicker.Value
					[DateTime]$EndHour = $EndTimePicker.Value
					
					
					#ConferenceDateTime, Caller, FeedbackText, Rating, TokenDescription, TokenValue	
					$CurrentDateTime = $null
					$CurrentCaller = $null
					
					$PreviousCaller = $null
					$PreviousRating = $null
					$PreviousFeedbackText = $null
					$PreviousTokenDescriptionString = $null
					$Counter = 0
					
					[string]$TokenString = "Voice"

										
					foreach ($t in $tempstoreVoice)
					{
						$Counter++
						
						#[string]$TokenString = "Voice"
						
						#CHECK Date
						[DateTime]$ConferenceDateTime = $t.ConferenceDateTime
						
						if($TimeFormatCheckBox.Checked)
						{
							#Do nothing, already GMT
						}
						else
						{
							$ConferenceDateTime = [System.TimeZone]::CurrentTimeZone.ToLocalTime($ConferenceDateTime)
						}						
										
						[string]$Caller = $t.Caller
						[string]$Rating = $t.Rating
						[string]$TokenValue = $t.TokenValue
						[string]$FeedbackText = $t.FeedbackText
						[string]$TokenId = $t.TokenId
						[string]$TokenDescription = $TokenDef[$TokenId]  #$t.TokenDescription		

						#CHECK IF IT'S A VIDEO CALL
						[int]$returnedInt = 0
						[bool]$result = [int]::TryParse($TokenId, [ref]$returnedInt)
						if($result -and (($returnedInt -ge 20 -and $returnedInt -le 30) -or ($returnedInt -ge 200 -and $returnedInt -le 210)))
						{
							[string]$TokenString = "Video"
						}
						
						
						#Built the Token Description
						if($ConferenceDateTime -eq $CurrentDateTime -and $Caller -eq $CurrentCaller)
						{
							$PreviousCaller = $Caller
							$PreviousRating = $Rating
							$PreviousFeedbackText = $FeedbackText
							
							if($TokenValue -eq "1")
							{
								if($PreviousTokenDescriptionString -eq $null -and !($PreviousTokenDescriptionString -match $TokenDescription))
								{
									$PreviousTokenDescriptionString = "${TokenDescription}"
									
								}
								elseif(!($PreviousTokenDescriptionString -match $TokenDescription))
								{
									$PreviousTokenDescriptionString = "${PreviousTokenDescriptionString}, ${TokenDescription}"
								}
							}
						}
						else
						{
							#Add the user info
							if($CurrentDateTime -gt $StartHour)
							{
								if($CurrentDateTime -lt $EndHour)
								{
									if($AndAboveCheckBox.Checked)
									{
										#Check that the rating is higher than the one chosen
										if($PreviousRating -ge $RatingTextBox.Text -or $RatingTextBox.Text -eq "ALL")
										{
											if($PreviousCaller -match $UserTextBox.Text)
											{
												if($PreviousTokenDescriptionString -imatch $ReasonTextBox.Text)
												{
													if(($TokenString -eq "Voice" -and $VoiceCheckBox.Checked) -or ($TokenString -eq "Video" -and $VideoCheckBox.Checked) -or ($VideoCheckBox.Checked -and $VoiceCheckBox.Checked))
													{
														$dgv.Rows.Add( @($CurrentDateTime.ToString("yyyy/MM/dd HH:mm:ss"),$PreviousCaller,$PreviousRating,$TokenString,$PreviousTokenDescriptionString, $PreviousFeedbackText) )
													}
												}
											}
											else
											{
												#Write-Host "Filtered due to SIP URI"
											}
										}
										else
										{
											#Write-Host "Ignoring Record due to rating"
										}
									}
									elseif($AndBelowCheckBox.Checked)
									{
										#Check that the rating is higher than the one chosen
										if($PreviousRating -le $RatingTextBox.Text -or $RatingTextBox.Text -eq "ALL")
										{
											if($PreviousCaller -match $UserTextBox.Text)
											{
												if($PreviousTokenDescriptionString -imatch $ReasonTextBox.Text)
												{
													if(($TokenString -eq "Voice" -and $VoiceCheckBox.Checked) -or ($TokenString -eq "Video" -and $VideoCheckBox.Checked) -or ($VideoCheckBox.Checked -and $VoiceCheckBox.Checked))
													{
														$dgv.Rows.Add( @($CurrentDateTime.ToString("yyyy/MM/dd HH:mm:ss"),$PreviousCaller,$PreviousRating,$TokenString,$PreviousTokenDescriptionString, $PreviousFeedbackText) )
													}
												}
											}
											else
											{
												#Write-Host "Filtered due to SIP URI"
											}
										}
										else
										{
											#Write-Host "Ignoring Record due to rating"
										}
									}
									else
									{
										if($RatingTextBox.Text -eq $PreviousRating -or $RatingTextBox.Text -eq "ALL")
										{
											if($PreviousCaller -match $UserTextBox.Text)
											{
												if($PreviousTokenDescriptionString -imatch $ReasonTextBox.Text)
												{
													if(($TokenString -eq "Voice" -and $VoiceCheckBox.Checked) -or ($TokenString -eq "Video" -and $VideoCheckBox.Checked) -or ($VideoCheckBox.Checked -and $VoiceCheckBox.Checked))
													{
														$dgv.Rows.Add( @($CurrentDateTime.ToString("yyyy/MM/dd HH:mm:ss"),$PreviousCaller,$PreviousRating,$TokenString,$PreviousTokenDescriptionString, $PreviousFeedbackText) )
													}
												}
											}
											else
											{
												#Write-Host "Filtered due to SIP URI"
											}
										}
										else
										{
											#Write-Host "Ignoring Record due to rating"
										}
									}
								}
								else
								{
									#Write-Host "Ignoring Record date $ConferenceDateTime because of end hour"
								}
							}
							else
							{
								#Write-Host "Ignoring Record date $ConferenceDateTime because of start hour"
							}
						
							#RESET VALUES TO DEFAULT
							$PreviousTokenDescriptionString = $null
							$TokenString = "Voice"
							#UPDATE THE VALUES WITH THE NEW CALLER INFO
							$CurrentCaller = $Caller
							$CurrentDateTime = $ConferenceDateTime
						}
						
						#ON FIRST PASS PICK UP THE FIRST TOKEN
						if($TokenValue -eq "1")
						{
							if($PreviousTokenDescriptionString -eq $null) #-and $PreviousTokenDescriptionString -match $TokenDescription
							{
								$PreviousTokenDescriptionString = "${TokenDescription}"
							}
						}
						
						
						#If it's the last record then add it.
						if($tempstoreVoice.count -eq $Counter)
						{
							#Add the user info
							if($CurrentDateTime -gt $StartHour)
							{
								if($CurrentDateTime -lt $EndHour)
								{
									if($AndAboveCheckBox.Checked)
									{
										#Check that the rating is higher than the one chosen
										if($PreviousRating -ge $RatingTextBox.Text -or $RatingTextBox.Text -eq "ALL")
										{
											if($PreviousCaller -match $UserTextBox.Text)
											{
												if($PreviousTokenDescriptionString -imatch $ReasonTextBox.Text)
												{
													if(($TokenString -eq "Voice" -and $VoiceCheckBox.Checked) -or ($TokenString -eq "Video" -and $VideoCheckBox.Checked) -or ($VideoCheckBox.Checked -and $VoiceCheckBox.Checked))
													{
														$dgv.Rows.Add( @($CurrentDateTime.ToString("yyyy/MM/dd HH:mm:ss"),$PreviousCaller,$PreviousRating,$TokenString,$PreviousTokenDescriptionString, $PreviousFeedbackText) )
													}
												}
											}
											else
											{
												#Write-Host "Filtered due to SIP URI"
											}
										}
										
									}
									elseif($AndBelowCheckBox.Checked)
									{
										#Check that the rating is higher than the one chosen
										if($PreviousRating -le $RatingTextBox.Text -or $RatingTextBox.Text -eq "ALL")
										{
											if($PreviousCaller -match $UserTextBox.Text)
											{
												if($PreviousTokenDescriptionString -imatch $ReasonTextBox.Text)
												{
													if(($TokenString -eq "Voice" -and $VoiceCheckBox.Checked) -or ($TokenString -eq "Video" -and $VideoCheckBox.Checked) -or ($VideoCheckBox.Checked -and $VoiceCheckBox.Checked))
													{
														$dgv.Rows.Add( @($CurrentDateTime.ToString("yyyy/MM/dd HH:mm:ss"),$PreviousCaller,$PreviousRating,$TokenString,$PreviousTokenDescriptionString, $PreviousFeedbackText) )
													}
												}
											}
											else
											{
												#Write-Host "Filtered due to SIP URI"
											}
										}
										else
										{
											#Write-Host "Ignoring Record due to rating"
										}
									}
									else
									{
										if($RatingTextBox.Text -eq $PreviousRating -or $RatingTextBox.Text -eq "ALL")
										{
											if($PreviousCaller -match $UserTextBox.Text)
											{
												if($PreviousTokenDescriptionString -imatch $ReasonTextBox.Text)
												{
													if(($TokenString -eq "Voice" -and $VoiceCheckBox.Checked) -or ($TokenString -eq "Video" -and $VideoCheckBox.Checked) -or ($VideoCheckBox.Checked -and $VoiceCheckBox.Checked))
													{
														$dgv.Rows.Add( @($CurrentDateTime.ToString("yyyy/MM/dd HH:mm:ss"),$PreviousCaller,$PreviousRating,$TokenString,$PreviousTokenDescriptionString, $PreviousFeedbackText) )
													}
												}
											}
											else
											{
												#Write-Host "Filtered due to SIP URI"
											}
										}
										else
										{
											#Write-Host "Ignoring Record due to rating"
										}
									}
								}
								else
								{
									#Write-Host "Ignoring Record date $ConferenceDateTime because of end hour"
								}
							}
							else
							{
								#Write-Host "Ignoring Record date $ConferenceDateTime because of start hour"
							}
						}
					}
					
				}
				else
				{
					Write-Host "Error: The monitoring database is a lower level than Skype for Business. This tool can only be used on Skype for Business monitoring servers." -foreground "red"
				}
			}
		}
		
		#ROW NUMBERS
		$RowNumberLabel.Text = "Rows: " + $dgv.Rows.Count
		Write-Host "SETTING ROW NUMBER"
		
		$dgv.Sort($dgv.Columns[0], [System.ComponentModel.ListSortDirection]::Descending)
		$dgv.PerformLayout()
	}
	else
	{
		Write-Host "The start time is higher than the end time. Fix the start and end times." -foreground "red"
	}
}


function MessageView([string]$Message, [string]$WindowTitle, [string]$DefaultText)
{
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
    
	#Message Text Start Text box ============================================================
	$MessageTextBox = new-object System.Windows.Forms.textbox
	$MessageTextBox.location = new-object system.drawing.size(5,5)
	$MessageTextBox.size = new-object system.drawing.size(670,445)
	$MessageTextBox.Multiline = $True	
	$MessageTextBox.Wordwrap = $True
	$MessageTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
	$MessageTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Top
	$MessageTextBox.text = $Message  
	$MessageTextBox.tabIndex = 0
	$MessageTextBox.Select(0,0)

    # Create the form.
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = $WindowTitle
    $form.Size = New-Object System.Drawing.Size(700,500)
    $form.FormBorderStyle = 'Sizable'
    $form.StartPosition = "CenterScreen"
    $form.Topmost = $True
	$form.MinimumSize = New-Object System.Drawing.Size(200,200) 
	[byte[]]$WindowIcon = @(66, 77, 56, 3, 0, 0, 0, 0, 0, 0, 54, 0, 0, 0, 40, 0, 0, 0, 16, 0, 0, 0, 16, 0, 0, 0, 1, 0, 24, 0, 0, 0, 0, 0, 2, 3, 0, 0, 18, 11, 0, 0, 18, 11, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114,0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0,198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 205, 132, 32, 234, 202, 160,255, 255, 255, 244, 229, 208, 205, 132, 32, 202, 123, 16, 248, 238, 224, 198, 114, 0, 205, 132, 32, 234, 202, 160, 255,255, 255, 255, 255, 255, 244, 229, 208, 219, 167, 96, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 248, 238, 224, 198, 114, 0, 198, 114, 0, 223, 176, 112, 255, 255, 255, 219, 167, 96, 198, 114, 0, 198, 114, 0, 255, 255, 255, 255, 255, 255, 227, 185, 128, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 241, 220, 192, 198, 114, 0, 198,114, 0, 248, 238, 224, 255, 255, 255, 244, 229, 208, 198, 114, 0, 198, 114, 0, 255, 255, 255, 255, 255, 255, 227, 185, 128, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 241, 220, 192, 198, 114, 0, 216, 158, 80, 255, 255, 255, 255, 255, 255, 252, 247, 240, 209, 141, 48, 198, 114, 0, 255, 255, 255, 255, 255, 255, 227, 185, 128, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 241, 220, 192, 198, 114, 0, 241, 220, 192, 255, 255, 255, 252, 247, 240, 212, 149, 64, 234, 202, 160, 198, 114, 0, 255, 255, 255, 255, 255, 255, 227, 185, 128, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 241, 220, 192, 205, 132, 32, 255, 255, 255, 255, 255, 255, 227, 185, 128, 198, 114, 0, 248, 238, 224, 202, 123, 16, 255, 255, 255, 255, 255, 255, 227, 185, 128, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 241, 220, 192, 234, 202, 160, 255, 255, 255, 255, 255, 255, 205, 132, 32, 198, 114, 0, 223, 176, 112, 223, 176, 112, 255, 255, 255, 255, 255, 255, 227, 185, 128, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 244, 229, 208, 252, 247, 240, 255, 255, 255, 237, 211, 176, 198, 114, 0, 198, 114, 0, 202, 123, 16, 248, 238, 224, 255, 255, 255, 255, 255, 255, 227, 185, 128, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 212, 149, 64, 255, 255, 255, 255, 255, 255, 255, 255, 255, 212, 149, 64, 198, 114, 0, 198, 114, 0, 198, 114, 0, 234, 202, 160, 255, 255,255, 255, 255, 255, 241, 220, 192, 205, 132, 32, 198, 114, 0, 198, 114, 0, 205, 132, 32, 227, 185, 128, 227, 185, 128, 227, 185, 128, 227, 185, 128, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 205, 132, 32, 227, 185, 128, 227, 185,128, 227, 185, 128, 219, 167, 96, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 0, 0)
	$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
	$form.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
     
    # Add all of the controls to the form.
	$form.controls.add($MessageTextBox)

    # Initialize and show the form.
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() > $null   # Trash the text of the button that was clicked.

}
	

# Activate the form ============================================================
$objForm.Add_Shown({$objForm.Activate()})
[void] $objForm.ShowDialog()	


# SIG # Begin signature block
# MIIcWAYJKoZIhvcNAQcCoIIcSTCCHEUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUMizPMK/SI0Oh4/PBgE0wqdWx
# bmSggheHMIIFEDCCA/igAwIBAgIQBsCriv7g+QV/64ncHMA83zANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMB4XDTE4MDExMzAwMDAwMFoXDTE5MDQx
# ODEyMDAwMFowTTELMAkGA1UEBhMCQVUxEDAOBgNVBAcTB01pdGNoYW0xFTATBgNV
# BAoTDEphbWVzIEN1c3NlbjEVMBMGA1UEAxMMSmFtZXMgQ3Vzc2VuMIIBIjANBgkq
# hkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAukBaV5eP8/bHNonSdpgvTK/2iYj9XRl4
# VzpuJE1fK2sk0ZnjIidsaYXhFpL1LUbUlalxnO7cbWY5ok5bHg0vPx9p8IHHBH28
# xrisz7wcXTTjXMrOL+yynDJYUCMpKV5rMkBn5kJJlLUrY5kcT6Y0fa4HKmvLYAVC
# 6T83mvUvwVs0TlLqY5Dcm/eoVzSmv9Frn3A5WNKxElhhUL2W6LEHdikzCRltk0+e
# g6OXRSYHwulwL+HzcQ+83YEVp/YG9GM+v3Ra4UeuSWaOkt4FQI5JGMlKvhQ3wSu4
# 455xAyj56MTul2FQ1s+j2KI/bvJOMwzO86RDwUC+yZuhh8+IYVObpQIDAQABo4IB
# xTCCAcEwHwYDVR0jBBgwFoAUWsS5eyoKo6XqcQPAYPkt9mV1DlgwHQYDVR0OBBYE
# FHXxYdsGH8A4rhw89n7VPGve7xb5MA4GA1UdDwEB/wQEAwIHgDATBgNVHSUEDDAK
# BggrBgEFBQcDAzB3BgNVHR8EcDBuMDWgM6Axhi9odHRwOi8vY3JsMy5kaWdpY2Vy
# dC5jb20vc2hhMi1hc3N1cmVkLWNzLWcxLmNybDA1oDOgMYYvaHR0cDovL2NybDQu
# ZGlnaWNlcnQuY29tL3NoYTItYXNzdXJlZC1jcy1nMS5jcmwwTAYDVR0gBEUwQzA3
# BglghkgBhv1sAwEwKjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNlcnQu
# Y29tL0NQUzAIBgZngQwBBAEwgYQGCCsGAQUFBwEBBHgwdjAkBggrBgEFBQcwAYYY
# aHR0cDovL29jc3AuZGlnaWNlcnQuY29tME4GCCsGAQUFBzAChkJodHRwOi8vY2Fj
# ZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRTSEEyQXNzdXJlZElEQ29kZVNpZ25p
# bmdDQS5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsFAAOCAQEATz4Xu/3x
# ae3iTkPfYm7uEWpB16eV1Ig+8FMDg6CJ+465oidj2amAjD1n+MwekysJOmcWAiEg
# R7TcQKUpgy5QTTJGSsPm2rwwcBL0jye6hXgs5eD8szZEhdJOnl1txRsdhMtilV2I
# H7X1nQ6S/eRu4WneUUF3YIDreqFYGLIfAobafEEufP7pMk05zgO6lqBM97ee+roR
# eP12IG7CBokmhzoERIDdTjfNEbDtob3OKPKfao2K8MJ079CSoG+NnpieO4CSRQtu
# kaCfg4rK9iCFIksrHq+qSMMRobnVwZq5tDZrkQOjO+lBdL0XWF4nrBavCjs4DjBh
# JHz6nkyqXDNAuTCCBTAwggQYoAMCAQICEAQJGBtf1btmdVNDtW+VUAgwDQYJKoZI
# hvcNAQELBQAwZTELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZ
# MBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNz
# dXJlZCBJRCBSb290IENBMB4XDTEzMTAyMjEyMDAwMFoXDTI4MTAyMjEyMDAwMFow
# cjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQ
# d3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVk
# IElEIENvZGUgU2lnbmluZyBDQTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoC
# ggEBAPjTsxx/DhGvZ3cH0wsxSRnP0PtFmbE620T1f+Wondsy13Hqdp0FLreP+pJD
# wKX5idQ3Gde2qvCchqXYJawOeSg6funRZ9PG+yknx9N7I5TkkSOWkHeC+aGEI2YS
# VDNQdLEoJrskacLCUvIUZ4qJRdQtoaPpiCwgla4cSocI3wz14k1gGL6qxLKucDFm
# M3E+rHCiq85/6XzLkqHlOzEcz+ryCuRXu0q16XTmK/5sy350OTYNkO/ktU6kqepq
# CquE86xnTrXE94zRICUj6whkPlKWwfIPEvTFjg/BougsUfdzvL2FsWKDc0GCB+Q4
# i2pzINAPZHM8np+mM6n9Gd8lk9ECAwEAAaOCAc0wggHJMBIGA1UdEwEB/wQIMAYB
# Af8CAQAwDgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHkGCCsG
# AQUFBwEBBG0wazAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29t
# MEMGCCsGAQUFBzAChjdodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRBc3N1cmVkSURSb290Q0EuY3J0MIGBBgNVHR8EejB4MDqgOKA2hjRodHRwOi8v
# Y3JsNC5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMDqg
# OKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURS
# b290Q0EuY3JsME8GA1UdIARIMEYwOAYKYIZIAYb9bAACBDAqMCgGCCsGAQUFBwIB
# FhxodHRwczovL3d3dy5kaWdpY2VydC5jb20vQ1BTMAoGCGCGSAGG/WwDMB0GA1Ud
# DgQWBBRaxLl7KgqjpepxA8Bg+S32ZXUOWDAfBgNVHSMEGDAWgBRF66Kv9JLLgjEt
# UYunpyGd823IDzANBgkqhkiG9w0BAQsFAAOCAQEAPuwNWiSz8yLRFcgsfCUpdqgd
# XRwtOhrE7zBh134LYP3DPQ/Er4v97yrfIFU3sOH20ZJ1D1G0bqWOWuJeJIFOEKTu
# P3GOYw4TS63XX0R58zYUBor3nEZOXP+QsRsHDpEV+7qvtVHCjSSuJMbHJyqhKSga
# OnEoAjwukaPAJRHinBRHoXpoaK+bp1wgXNlxsQyPu6j4xRJon89Ay0BEpRPw5mQM
# JQhCMrI2iiQC/i9yfhzXSUWW6Fkd6fp0ZGuy62ZD2rOwjNXpDd32ASDOmTFjPQga
# GLOBm0/GkxAG/AeB+ova+YJJ92JuoVP6EpQYhS6SkepobEQysmah5xikmmRR7zCC
# BmowggVSoAMCAQICEAMBmgI6/1ixa9bV6uYX8GYwDQYJKoZIhvcNAQEFBQAwYjEL
# MAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3
# LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgQXNzdXJlZCBJRCBDQS0x
# MB4XDTE0MTAyMjAwMDAwMFoXDTI0MTAyMjAwMDAwMFowRzELMAkGA1UEBhMCVVMx
# ETAPBgNVBAoTCERpZ2lDZXJ0MSUwIwYDVQQDExxEaWdpQ2VydCBUaW1lc3RhbXAg
# UmVzcG9uZGVyMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAo2Rd/Hyz
# 4II14OD2xirmSXU7zG7gU6mfH2RZ5nxrf2uMnVX4kuOe1VpjWwJJUNmDzm9m7t3L
# helfpfnUh3SIRDsZyeX1kZ/GFDmsJOqoSyyRicxeKPRktlC39RKzc5YKZ6O+YZ+u
# 8/0SeHUOplsU/UUjjoZEVX0YhgWMVYd5SEb3yg6Np95OX+Koti1ZAmGIYXIYaLm4
# fO7m5zQvMXeBMB+7NgGN7yfj95rwTDFkjePr+hmHqH7P7IwMNlt6wXq4eMfJBi5G
# EMiN6ARg27xzdPpO2P6qQPGyznBGg+naQKFZOtkVCVeZVjCT88lhzNAIzGvsYkKR
# rALA76TwiRGPdwIDAQABo4IDNTCCAzEwDgYDVR0PAQH/BAQDAgeAMAwGA1UdEwEB
# /wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwggG/BgNVHSAEggG2MIIBsjCC
# AaEGCWCGSAGG/WwHATCCAZIwKAYIKwYBBQUHAgEWHGh0dHBzOi8vd3d3LmRpZ2lj
# ZXJ0LmNvbS9DUFMwggFkBggrBgEFBQcCAjCCAVYeggFSAEEAbgB5ACAAdQBzAGUA
# IABvAGYAIAB0AGgAaQBzACAAQwBlAHIAdABpAGYAaQBjAGEAdABlACAAYwBvAG4A
# cwB0AGkAdAB1AHQAZQBzACAAYQBjAGMAZQBwAHQAYQBuAGMAZQAgAG8AZgAgAHQA
# aABlACAARABpAGcAaQBDAGUAcgB0ACAAQwBQAC8AQwBQAFMAIABhAG4AZAAgAHQA
# aABlACAAUgBlAGwAeQBpAG4AZwAgAFAAYQByAHQAeQAgAEEAZwByAGUAZQBtAGUA
# bgB0ACAAdwBoAGkAYwBoACAAbABpAG0AaQB0ACAAbABpAGEAYgBpAGwAaQB0AHkA
# IABhAG4AZAAgAGEAcgBlACAAaQBuAGMAbwByAHAAbwByAGEAdABlAGQAIABoAGUA
# cgBlAGkAbgAgAGIAeQAgAHIAZQBmAGUAcgBlAG4AYwBlAC4wCwYJYIZIAYb9bAMV
# MB8GA1UdIwQYMBaAFBUAEisTmLKZB+0e36K+Vw0rZwLNMB0GA1UdDgQWBBRhWk0k
# tkkynUoqeRqDS/QeicHKfTB9BgNVHR8EdjB0MDigNqA0hjJodHRwOi8vY3JsMy5k
# aWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURDQS0xLmNybDA4oDagNIYyaHR0
# cDovL2NybDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEQ0EtMS5jcmww
# dwYIKwYBBQUHAQEEazBpMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2Vy
# dC5jb20wQQYIKwYBBQUHMAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9E
# aWdpQ2VydEFzc3VyZWRJRENBLTEuY3J0MA0GCSqGSIb3DQEBBQUAA4IBAQCdJX4b
# M02yJoFcm4bOIyAPgIfliP//sdRqLDHtOhcZcRfNqRu8WhY5AJ3jbITkWkD73gYB
# jDf6m7GdJH7+IKRXrVu3mrBgJuppVyFdNC8fcbCDlBkFazWQEKB7l8f2P+fiEUGm
# vWLZ8Cc9OB0obzpSCfDscGLTYkuw4HOmksDTjjHYL+NtFxMG7uQDthSr849Dp3Gd
# Id0UyhVdkkHa+Q+B0Zl0DSbEDn8btfWg8cZ3BigV6diT5VUW8LsKqxzbXEgnZsij
# iwoc5ZXarsQuWaBh3drzbaJh6YoLbewSGL33VVRAA5Ira8JRwgpIr7DUbuD0FAo6
# G+OPPcqvao173NhEMIIGzTCCBbWgAwIBAgIQBv35A5YDreoACus/J7u6GzANBgkq
# hkiG9w0BAQUFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5j
# MRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBB
# c3N1cmVkIElEIFJvb3QgQ0EwHhcNMDYxMTEwMDAwMDAwWhcNMjExMTEwMDAwMDAw
# WjBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQL
# ExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBBc3N1cmVkIElE
# IENBLTEwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDogi2Z+crCQpWl
# gHNAcNKeVlRcqcTSQQaPyTP8TUWRXIGf7Syc+BZZ3561JBXCmLm0d0ncicQK2q/L
# XmvtrbBxMevPOkAMRk2T7It6NggDqww0/hhJgv7HxzFIgHweog+SDlDJxofrNj/Y
# MMP/pvf7os1vcyP+rFYFkPAyIRaJxnCI+QWXfaPHQ90C6Ds97bFBo+0/vtuVSMTu
# HrPyvAwrmdDGXRJCgeGDboJzPyZLFJCuWWYKxI2+0s4Grq2Eb0iEm09AufFM8q+Y
# +/bOQF1c9qjxL6/siSLyaxhlscFzrdfx2M8eCnRcQrhofrfVdwonVnwPYqQ/MhRg
# lf0HBKIJAgMBAAGjggN6MIIDdjAOBgNVHQ8BAf8EBAMCAYYwOwYDVR0lBDQwMgYI
# KwYBBQUHAwEGCCsGAQUFBwMCBggrBgEFBQcDAwYIKwYBBQUHAwQGCCsGAQUFBwMI
# MIIB0gYDVR0gBIIByTCCAcUwggG0BgpghkgBhv1sAAEEMIIBpDA6BggrBgEFBQcC
# ARYuaHR0cDovL3d3dy5kaWdpY2VydC5jb20vc3NsLWNwcy1yZXBvc2l0b3J5Lmh0
# bTCCAWQGCCsGAQUFBwICMIIBVh6CAVIAQQBuAHkAIAB1AHMAZQAgAG8AZgAgAHQA
# aABpAHMAIABDAGUAcgB0AGkAZgBpAGMAYQB0AGUAIABjAG8AbgBzAHQAaQB0AHUA
# dABlAHMAIABhAGMAYwBlAHAAdABhAG4AYwBlACAAbwBmACAAdABoAGUAIABEAGkA
# ZwBpAEMAZQByAHQAIABDAFAALwBDAFAAUwAgAGEAbgBkACAAdABoAGUAIABSAGUA
# bAB5AGkAbgBnACAAUABhAHIAdAB5ACAAQQBnAHIAZQBlAG0AZQBuAHQAIAB3AGgA
# aQBjAGgAIABsAGkAbQBpAHQAIABsAGkAYQBiAGkAbABpAHQAeQAgAGEAbgBkACAA
# YQByAGUAIABpAG4AYwBvAHIAcABvAHIAYQB0AGUAZAAgAGgAZQByAGUAaQBuACAA
# YgB5ACAAcgBlAGYAZQByAGUAbgBjAGUALjALBglghkgBhv1sAxUwEgYDVR0TAQH/
# BAgwBgEB/wIBADB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGGGGh0dHA6Ly9v
# Y3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2NhY2VydHMuZGln
# aWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDCBgQYDVR0fBHow
# eDA6oDigNoY0aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJl
# ZElEUm9vdENBLmNybDA6oDigNoY0aHR0cDovL2NybDQuZGlnaWNlcnQuY29tL0Rp
# Z2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDAdBgNVHQ4EFgQUFQASKxOYspkH7R7f
# or5XDStnAs0wHwYDVR0jBBgwFoAUReuir/SSy4IxLVGLp6chnfNtyA8wDQYJKoZI
# hvcNAQEFBQADggEBAEZQPsm3KCSnOB22WymvUs9S6TFHq1Zce9UNC0Gz7+x1H3Q4
# 8rJcYaKclcNQ5IK5I9G6OoZyrTh4rHVdFxc0ckeFlFbR67s2hHfMJKXzBBlVqefj
# 56tizfuLLZDCwNK1lL1eT7EF0g49GqkUW6aGMWKoqDPkmzmnxPXOHXh2lCVz5Cqr
# z5x2S+1fwksW5EtwTACJHvzFebxMElf+X+EevAJdqP77BzhPDcZdkbkPZ0XN1oPt
# 55INjbFpjE/7WeAjD9KqrgB87pxCDs+R1ye3Fu4Pw718CqDuLAhVhSK46xgaTfwq
# Ia1JMYNHlXdx3LEbS0scEJx3FMGdTy9alQgpECYxggQ7MIIENwIBATCBhjByMQsw
# CQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cu
# ZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQg
# Q29kZSBTaWduaW5nIENBAhAGwKuK/uD5BX/ridwcwDzfMAkGBSsOAwIaBQCgeDAY
# BgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3
# AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEW
# BBQnL4MzlIkpR/IbsP/QbQNV6YzWZTANBgkqhkiG9w0BAQEFAASCAQB8EZJjXa0F
# ALAyBDoXMMz6812XZsoEK2wTbFbKDA3HH757Pl7pqVPy/D3f0Ygq4uiGhgoq/bvK
# VE3ue/zpjW4gld1/n81QC0Nrh3/pkaw1ehH3q/1ej8ytJi6paSiTsvySUXBudnNf
# gmKRYj5d8S1viVu4tTpFvIGme+1E0PSGh20hqGhKyI/wNAcSWpTZcGrLQOBgsjxS
# CM+fD1fMah6Zgzyt0DzzZfM1citbL7taq6rQ3OtRrDeTU5KSD5jPEeQVTq5vL8Zc
# pxr2JIxAgUm3SLFB9K7YfFuphRRC5CPc2P7SQ4IREdkKqLIAjQmsHI0aJN8nu9XW
# lckAk+LHrXwjoYICDzCCAgsGCSqGSIb3DQEJBjGCAfwwggH4AgEBMHYwYjELMAkG
# A1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRp
# Z2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgQXNzdXJlZCBJRCBDQS0xAhAD
# AZoCOv9YsWvW1ermF/BmMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZI
# hvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xODAzMTYwOTUxNDFaMCMGCSqGSIb3DQEJ
# BDEWBBRHIO5cEuHBDJjpEzaoYzVaY+/tdzANBgkqhkiG9w0BAQEFAASCAQBfRGBS
# Cl9veFL20A5a9EqEBmOmYxe/BDTgEmInnAY3PUID3i/LZucpiFUBIOKrdtSlIJag
# U+ZXCNQWIpOjMTjjyDQR7+2Blhn42ZGPgzTJmFn5Vsy5AEoI+h+/KLUGmjXetLmf
# +v9zRsFn0BH89+kH9t71t10GfCuVC2P+5sSPyaHj44NAyRKNob/muWliKoKJEG9Y
# uVCoTp4ggTWD8yVGYHDELHOUqBYreHrZsWquYNEtEKhtA9cos4iNqDZaKhjtxaRn
# sGw+yn6cc/onEzQKJGmYtaEZtKsgyDJnT0xPWS43in4zWXQEovOjMysmLMfwYbGN
# TA94c59szFwsClFL
# SIG # End signature block
