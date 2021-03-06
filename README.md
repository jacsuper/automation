# automation
Useful  Automation Scripts

Often gathered from all over the internet (for example thank you Mr Excel!), self written, customized, tweaked etc. In other words, licenses and copied from stackoverflows or sites author names etc. could have been lost: apologies. Let me know where attribution needed.


/vba - Scripts mostly to automate Microsoft Windows version of Outlook

/applescript - Scripts mostly to automate MacOS and MacOS version of Outlook

---------------------------
*Applescript:*

ArchiveMails.applescript - Move Outlook selected emails to Archive

JoinNextMeeting.applescript - Find Outlook Calendar next meeting (or current started 5 mins or less ago), open video conference URL in Chrome.  Uses getNextMeeting gist from https://gist.github.com/ivoleitao/f3937582744c412c1042c8d4617ea0d8 Author: Ivo Leitão - Currently Outlook has a bug where not all meetings are returned :( .

NewNoteOutlook.applescript - Create new text empty email, addressed to self, tab into Subject field

ParseTasksFromEmail.applescript - Parse content text from selected mails, create new emails for each line starting with *

RaiseVideoCallTab.aspplescript - Find active video call tab in Chrome, bring Chrome and Tab to front

SaveAttachments.applescript - Save Attachments from selected emails in Outlook to a folder

ZipFilesWithPassword.applescript - Zip and Encrypt selected files from finder into a zip with prompted password

OpenWorkplaceChatUnread.applescript - Open Workplace, open Chat, then click Unread (uncomment keystroke to open/focus search)

RunApplescriptAutomator.txt - How to run an AppleScript from an Automator App - to be able to have an applescript in Dock and shortcut key

RunWithStreamDeck.txt - Link to working Streamdeck AppleScript integration - to bind scripts to StreamDeck keys

*VBA Script:*

outlookVBACombined.vba - combined VBA file with all other VBA files for ease of install

archiveAndToDone.vba - clear categories from selected items and move to an Archive folder

categoryToAppointsments.vba - on incoming or saved appointment, set category according to subject or ownership

clearCategoriesAndFlags.vba - clear flag and categories from selected items

clearDeletedItems.vba - clear deleted items folder

clearJunkFolders.vba - move set of junk folders contents to deleted items

confidential.vba - set confidential flag and ad to subject Confidential on currently selcted item

delaySend.vba - send email in 5 minutes (allowing delete from outbox if mistake made) - safer alternative than Send

deleteAttachmentsSaveToFolder.vba - to minimize mailbox size, for selected items save attachments to selcted folder, then remove from items

deleteSelected.vba - selete selected items

duplicateMail.vba - duplicate the email X times into New emails. Useful with simpleTemplates.

simpleTemplate.vba - Idea is to save emails to a template email folder.  In said emails, templatize with {PLACEHOLDER,defaultvalue}.  Also support set of defaults like TO and CC to fill in mail merge like Name from TO email field.  Running macro fills in defaults, or asks user to accept defaultvalue or type in a textbox parameters to search and replace.  This gives a quick way to fill and replace 1 template, then use DuplicateMail, then fill parameters and send.

expandFolders.vba - Expand all folders in Outlook folder nav

fileEmailToFolder.vba - In a low tech way try and detect where to file an email to.  Start with "meeting notes" folders, then fall through to "project folders". If no luck - suggest creating a new meeting or project folder and move.  The filer will fuzzy match words from the subject of the email with words in folder mails.  Asks yes/no to the user.  If no, searches on. 

findContactActivities.vba - Find all emails related to a contact (from selected item). Author:  Victor Beekman victor[dot]beekman"at"xs4all{dot}nl 

mailSaveAttachments.vba - On new email arriving, save attachments to a system folder

newMail.vba - Open new mail dialog to self, seting focus to the body ready for typing.

openAppointmentCopy.vba - Outlook macro to create a new appointment with specific details of the currently selected appointment and show it in a new window. Author: Robert Sparnaaij, https://www.howto-outlook.com/howto/openapptcopy.htm

quickNav.vba - Examples to bind macros to toolbar butons that open a specific Outlook mail folder. Also JumpFolder fuzzy match open example (type word, open first folder with that word in the name).

replyWithAttachments.vba - Reply and Reply all WITH attachments (versus Outlook behaviour of not including atachments).

splitActionsFromMail.vba - Create new emails from lines in body of selected emails that start with * and send to self (using email as tasks).

viewHeaders.vba - Easy view email headers - Author: BlueDevilFan techniclee.wordpress.com


---------------------------

 -- Jac --
 
 License
Copyright 2017 © Jac Steyn

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
