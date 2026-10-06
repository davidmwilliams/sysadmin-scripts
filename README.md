# sysadmin-scripts
Useful scripts for systems administrators

## Active Directory ##
* CreateNewServerADGroups.ps1 - create two new AD groups for a server, to control who has remote access to it, and who has admin access on it
* UserLogons.vbs - reports on who didn't log in today, across all domain controllers
* WhyUserLockedOut.ps1 - seeks to identify why a user's account is locked.

## Exchange / Office 365 ##
* WhatMailboxesForwardWhere.ps1 - check where all your O365 mailboxes are being forwarded to

## Networking ##
* IsPortOpen.ps1 - determine if a port is open on a particular remote server

## Miscellaneous ##
* GetRandomPassword.ps1 - generates a random strong password
* Set-WindowsSleepSettings.ps1 - Configures the sleep settings for the current power plan, defaults to sleeping after an hour idle on battery and never sleeping when plugged in.
* WhoShutItDown.ps1 - identifies who shut the server down.
