#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


;*********************************************************************
; AUTHOR:
;	CLARK MULVEY	
; NAME:	
;	AUTHORIZE
; DESCRIPTION:
; 	Implementation of all authorize functions
; FUNCTIONS:
; 	shouldAuthorize()
;	hasBeenAuthorized()
;	isHoldTypeAuthorizable()
;	isRateCodeAuthorizable()
;	authorize()
;	releaseAuthorization()
;	changeToNotGtd()
; NOTES:
;	All notes are on individual functions
;	

;-------------shouldAuthorize()-----------------------
;	This function essentially just calls other functions to see if 
;	The reservation is authorizable - what it's lookin for is 
; 	the correct rate code, the correct hold type, and if the reservation
;	has already been authorized or not.
;	If the rate code and hold type are valid, then the function will go
;	into the payments window under the folio tab, where it will then check
;	to see if the credit card has already been authorized. If it has not,
;	function returns true, if none or some of those conditions are not met,
;	the function returns false
;-------------------------------------------------------
shouldAuthorize()
{
	if (isRateCodeAuthorizable() && isHoldTypeAuthorizable())
	{
		; If the rate code and hold type are authorizable, go into payments
		;	window
		; Click Folio
		Click, 	830, 150, 1		
		Sleep,	7500	

		; Click Payments
		Click, 	410, 350, 1		
		Sleep,  2000	
		
		if (!hasBeenAuthorized())
		{
			return true
		}
		else
		{
			; Click Payments cancel button
			Click,	 1300, 675, 1
			return false
		}
	}
	else 
	{
		return false
	}
}

;-----------------hasBeenAuthorize()----------------------------
;	This function assumes the payments window is already open
;	This function will check the "authorized" box in the reservation
;	to see if it's none equal to the string "$0.00" - if it's not
;	function returns true, if it is, returns false
;	There's a couple of repeat clicks here that serve the same exact purpose
;	For example - the "click magnifying glass" bit
;	This is because sometimes the windows will load differently based on
;	if there's a credit card or not, if they've been
;	authorized or not, etc. The multiple clicks are used to be universal
;	in every condition they will open the correct window, and any "extra"
;	clicks, although they use time, don't actually do anything to the reservation
;	so they are harmless
;---------------------------------------------------------------
hasBeenAuthorized()
{	
	; Clear the clipboard so we aren't working with previous copied information
	Clipboard := ""
	
	; Click Magnifying glass
	Click, 	1300, 490, 1	
	Sleep, 	2000
	
	; Click Magnifying glass
	Click, 	1350, 490, 1	
	Sleep, 	2000
	
	; Click 'OK' box on "Choose folio to authorize" box, to get out of it
	;	This is in the specific case that there is a good hold type and
	;	There is a good rate code, but no credit card is on file
	Click, 	970, 605, 1
	Sleep, 	2000
	
	; Double Click amount authorized box
	Click,	1380, 520, 2	
	Sleep, 	2000
	Send, 	{ctrldown}c{ctrlup}
	Sleep, 	100
	
	if (Clipboard = "")
	{
		; Double Click amount authorized box
		Click,	1410, 520, 2	
		Sleep, 	2000
		Send, 	{ctrldown}c{ctrlup}
		Sleep, 	100
	}
	
	if (Clipboard = "")
	{
		; Double Click amount authorized box
		Click,	1440, 520, 2	
		Sleep,	2000
		Send, 	{ctrldown}c{ctrlup}
		Sleep, 	100
	}
	
	if (Clipboard = "")
	{
		; Double Click amount authorized box
		Click,	1500, 520, 2	
		Sleep, 	2000
		Send, 	{ctrldown}c{ctrlup}
		Sleep, 	100
	} 
	
	; Close amount authorized window
	Click, 995, 790	
	Sleep, 3000				
	
	if (clipboard = "$0.00")
	{
		return false
	}
	else
	{
		return true
	}
}

;-----------------isHoldTypeAuthorizable()----------------------------
;	Will select the Hold Type box and copy it to clipboard
;	If the hold type is "CC" or "FULLPAY" return true, else return false
;---------------------------------------------------------------
isHoldTypeAuthorizable()
{
	; Clear the clipboard so we aren't working with previous copied information
	Clipboard := ""
	
	Sleep,	250
	
	; Get hold type
	Click, 	1180, 535			; Click ID number Key box
	loop, 	2					; Tab twice to get to hold Type box
	{
		Send, 	{Tab}
		Sleep, 	100
	}
	
	Send, 	{ctrldown}c{ctrlup}	; Copy hold type to clipboard
	Sleep, 	100	
	
	if (Clipboard = "CC" or Clipboard = "FULLPAY")
	{
		return true
	}
	else
	{
		return false
	}
}

;-----------------isRateCodeAuthorizable()----------------------------
;	Will select the rate code box, copy it to the clipboard, and 
;	compare it against a list of authorizable rate codes, if it is equal
; 	to one of them return true, else return false
;	This is a really ugly if statement - but unfortunatley AutoHotkey only
;	lets you put this on one line - I have added a list of the rate codes in
;	a comment block above the if statement to track what is in that statement
;	Should note also that rates change - so if we need to add or subtract a 
;	rate code from this list, just change the comment box and either add or
;	delete a boolean within the if statement
;---------------------------------------------------------------
isRateCodeAuthorizable()
{
	; Clear clipboard
	Clipboard := ""
	
	; Get rate code
	Click, 	350, 320, 2			; Double click the Rate Code
	Sleep, 	2000
	Send, 	{ctrldown}c{ctrlup}	; Copy rate code to clipboard
	Sleep, 	100
	
	;check the following rates:
	; XW
	; ATFC
	; X1
	; SIEN
	; SHL
	; BW
	; AAA
	; AARP
	; LOC
	; SP
	; MINE
	; SELOCAL
	; OWNER
	; 2UB
	; GOVT
	; OMEGA
	; XN
	; EX1
	; 2U
	; GRP2
	; RACK
	; SCHOOL
	; SEARHC
	; SERRC
	; AMHS
	; BK1
	; CRON
	; EP
	; FEDN
	; FDRN
	; FF
	; IT
	; ITB
	; KROC
	; KRON
	; NVLC
	; OI
	; SOAC
	; BV
	; SEAK

	if (Clipboard = "XW" or Clipboard = "ATFC" or Clipboard = "X1" or Clipboard = "THL" or Clipboard = "SIEN" or Clipboard = "BW" or Clipboard = "AAA" or Clipboard = "AARP" or Clipboard = "LOC" or Clipboard = "SP" or Clipboard = "MINE" or Clipboard = "SELOCAL" or Clipboard = "OWNER" or Clipboard = "2UB" or Clipboard = "GOVT" or Clipboard = "OMEGA" or Clipboard = "XN" or Clipboard = "EX1" or Clipboard = "2U" or Clipboard = "GRP2" or Clipboard = "RACK" or Clipboard = "SCHOOL" or Clipboard = "SEARHC" or Clipboard = "SERRC" or Clipboard = "AMHS" or Clipboard = "BK1" or Clipboard = "CRON" or Clipboard = "EP" or Clipboard = "FEDN" or Clipboard = "FDRN" or Clipboard = "FF" or Clipboard = "IT" or Clipboard = "ITB" or Clipboard = "KROC" or Clipboard = "KRON" or Clipboard = "NVLC" or Clipboard = "OI" or Clipboard = "SOAC" or Clipboard = "BV" or Clipboard = "SEAK")
	{
		return true
	}
	else
	{
		return false
	}
}

;-----------------authorize()----------------------------
;	Authorizes the credit card for amount specified - in this function
;	I just hardcoded a value of 1 - but this can be changed, and can
;	even be added as a parameter to make it more dynamic
;---------------------------------------------------------------
authorize()
{
	authAmount := 1
	
	Click, 	1275, 490, 1	; Highlight bar in Payments Popup
	Click,  1390, 490, 1	; Click Authorize button
	Sleep,	5000
	Click, 	1160, 580, 2	; Double click Request Auth amount text box
							; 	We double click so we can overwrite 
							;  	with next command
	Sleep, 	2000
	Send,   % authAmount	; Send authAmount to Req Auth Amount text box
	Sleep, 	2000
	Click, 	925, 695		; Click OK button
	Sleep, 	2000
	Click, 	1015, 645		; Click "NO" in reference to do we want to swipe
							; 	credit card
	Sleep, 	20000			; wait 20 seconds for the authorization
	
	Click, 	970, 605, 1		; Click "OK" when the credit card declines
	Sleep, 	2000			;	Note this doesn't always happen - but it always clicks
							;	Just in case it does happen
	;Click, 1305, 675, 1	; Click "Close" on the payments window
	;Sleep, 1000
	return
}

;-----------------authorize()----------------------------
;	Releases the Authorization so that the authorization doesn't come up
;	as a pending charge on the guest's bank statement
;	There's a couple of repeat clicks here that serve the same exact purpose
;	For example - the "click magnifying glass" bit
;	This is because sometimes the windows will load differently based on
;	if there's a credit card or not, if they've been
;	authorized or not, etc. The multiple clicks are used to be universal
;	in every condition they will open the correct window, and any "extra"
;	clicks, although they use time, don't actually do anything to the reservation
;	so they are harmless
;---------------------------------------------------------------
releaseAuthorization()
{
	; Click Magnifying button to open authorized window
	Click, 	1300, 490, 1	
	Sleep,	 2000
	
	; Click Magnifying button to open authorized window
	Click, 	1350, 490, 1	
	Sleep, 	2000
	
	; Click "Release Authorization" button
	Click, 	1200, 520, 1
	Sleep, 	6000
	
	; Click "OK" button to release authorization
	Click,	925, 625, 1
	Sleep, 	10000
	
	; Click "OK" button to get rid of credit card release pop up box
	Click, 	965, 605, 1
	Sleep, 	5000
	
	; Click "Close" button to close out of authorization window
	Click, 	995, 790, 1
	Sleep, 	3000				
	
	; Click "Close" on the payments window
	Click, 1305, 675, 1		
	Sleep, 3000
	return
}

;-----------------changeToNotGtd()----------------------------
;	This is the whole point of this entire script - to find out which cards decline,
;	and then to flag them as not guarenteed for the front desk to find and call
;	Closes out of the payments window, and then changes into the guest reservation
;	Tab, after which it will change the hold type to 6pm hold, put a note in the comments
;	about the credit card declining, and then save the reservation
;	Should note that reservations cannot be saved without all necessary guest info
;	text boxes having something in them - often online reservations will leave these blank
;	This will put a "-" in most of the boxes, and put the state to "AK" and the 
; 	country to "United States"
;---------------------------------------------------------------
changeToNotGtd()
{
	; Click "Close" on the payments window
	Click, 1305, 675, 1		
	Sleep, 1000
	
	; Click on the Guest Info Tab
	Click, 775, 145, 1
	Sleep, 1000

	; Click on ID Number Key box, then once to get to Comments text box
	Click, 1180, 535	
	Sleep, 250
	Send, {Tab}
	Sleep, 250
	
	; Put message in comments box, then tab to hold type box
	Send, cc declined flagged automatically
	Sleep, 250
	Send, {Enter}
	Sleep, 250
	Send, {Tab}
	Sleep, 250
	
	; Press UP seven times to make sure we are at the top of the list
	loop, 7
	{
		Send, {Up}
		Sleep, 100
	}
	
	; Press DOWN twice to make sure we are on 6PM hold
	loop, 2					
	{
		Send, {Down}
		Sleep, 100
	}
	
	; This is ugly, but there's no other way to fill out these boxes as there 
	; 	are inconsistent tab amounts between the different text boxes
	; Make sure the reservation info is filled out
	; Click Address box
	Click, 515, 430, 1
	Sleep, 150
	Send, -
	Sleep, 150
	Send, {Tab}
	Sleep, 150
	Send, -
	Sleep, 150
	Send, {Tab}
	Sleep, 150
	Send, AK
	Sleep, 150
	Send, {Tab}
	Sleep, 150
	Send, -
	Sleep, 150
	Send, {Tab}
	Sleep, 150
	Send, {Tab}
	Sleep, 150
	Send, United States
	Sleep, 150
	Send, {Tab}
	Sleep, 150
	Send, {Tab}
	Sleep, 150
	Send, -
	Sleep, 150
	
	
	; Click Save
	Click, 280, 135, 1
	Sleep, 5000
	
	; Before going back to reservations page, these need to be closed out again
	; click runtriz box
	Click, 1025, 690, 1
	Sleep, 1000

	; click rewards points box
	Click, 960, 600, 1
	Sleep, 1000
	
	return
}
