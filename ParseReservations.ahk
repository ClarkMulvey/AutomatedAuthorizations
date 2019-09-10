#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


#include Authorize.ahk

;*********************************************************************
; AUTHOR:
;	CLARK MULVEY	
; NAME:	
;	Parse Reservations
; DESCRIPTION:
; 	Open all of the incoming reservations one at a time, perform
;	some action, and then close to reopen the next one
; HOTKEY(S);
; 	CTRL + P (^p)
; FUNCTIONS:
;	enterReservation()
; NOTES:
;	Everything in this and accompanied files assume we are using the Jonas Chorum
;	Property Management System for hotels
;	This script will need the page scrolled all the way to the top before
;  	Running, it will automatically find declined credit cards, and put them
; 	on a 6PM hold. There have been a few discrepancies, if any are found
;	please report them to me - clarkgmulvey3@gmail.com and I will fix them
;	Also, currently there is no way to pause the script, only a way to stop it.
;	Press ESC to stop the script mid run.
;	Before running this script the number of reservations needs to be manually put 
;	into the code. This is the numReservations variable Line - 49 
;	If you want to stop mid script and then pick up where you left off, just put in
;	the reservation number you left off into the currentReservation variable

; ---------------EnterReservation()----------------
;	Simple function to enter the reservation, and clear out any pop up boxes
;	that appear when the reservation automatically loads.
;----------------------------------------------------
enterReservation()
{
	; enter and wait to load
	Send, 	{Enter}
	Sleep, 	10000

	; click runtriz box
	Click, 	1025, 690, 1
	Sleep, 	1000

	; click rewards points box
	Click, 	960, 600, 1
	Sleep, 	1000
	return
}

^p::

numReservations 	:= 3
currentReservation 	:= 0

; -----------------Main Loop------------------
loop, %numReservations%
{
	;---------------PUT CURSOR ON RESERVATION-------------
	; Start by clicking the "Folio" tab twice, this orients the reservations page
	; 	So we never run into duplicate reservations (The "other names")
	; 	are at the bottom of the page
	Click, 	1480, 170, 1
	Sleep,	1500
	Click, 	1480, 170, 1
	Sleep, 	1500
	
	; Always start parsing by clicking in the starting box (This orients us so we can 
	;	run the rest of the find reservation portion of the script
	Click, 	460, 225, 1
	sleep, 	150
	
	; If it's the first reservation, just tab over twice, otherwise start looping through
	;	tabbing 26 times per loop to get to the next line
	if (currentReservation = 0)
	{
		Send, {Tab 2}
	}
	else
	{	
		loop, %currentReservation%
		{
			send, 	{Tab 26}
			
			; Need a sleep in between parsing the lines, otherwise it bugs out
			Sleep, 	250
		}
		
		; Once the desired reservation's line has been met, tab over 6 times to highlight
		;	the last name of guest
		Send, {Tab 6}
		
	}
	
	; --------------Enter Reservation---------------
	enterReservation()

	;---------------Authorize--------------------
	; Will check to see if the reservation is one we want to authorize, if it is
	;	authorize it - if it doesn't authorize, change the hold type of the reservation
	;	to 6PM hold - this will make it very easy for front desk to find to contact the 
	; 	guest with the declined credit card
	if (shouldAuthorize())
	{
		authorize()
		
		;---------------Check if Authorized------------
		; After the authorize script has been run, check again to see if the authorization
		;	amount changed - if it did not, then there is a bad credit card on file
		;	if it did change - release the authorization so it doesn't come up on the 
		;	guest's bank statement
		if (hasBeenAuthorized())
		{
			releaseAuthorization()
		}
		else
		{
			changeToNotGtd()
		}
	}
	
	;---------------Leave Reservation------------
	Sleep, 	2000
	Send, 	!{Left}	; Hotkey for Browser Back button
	Sleep, 	10000
	currentReservation++
}
return

Escape::
ExitApp
return