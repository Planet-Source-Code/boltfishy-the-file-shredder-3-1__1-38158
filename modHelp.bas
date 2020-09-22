Attribute VB_Name = "HELP"
' App Name                : The File Shredder
' App Version             : 2.50
' Purpose                 : Securely delete files
' Written by              : Mischa Balen
' Contact E-mail          : boltfish@eml.cc
' Website                 : www26.brinkster.com/boltfish/
' Date Created            : July / August 2002
' Last Updated            : 21st August 2002
'----------------------------------------------------------

'CONTENTS:

'   1. Foreword
'   2. Global Variables
'   3. Application Model and Functions
'   4. Delete Clicked
'   5. Shred File Function
'   6. Binary Function
'   7. Adding Context Menus
'   8. Custom Deletion Methods


'           Brief Foreword and Introduction
'      ----------------------------------------

'Hello all! Thanks for downloading my code. If you have
'downloaded it from Planet Source Code, then I'm not
'going to beg for 5 globes (it doesn't deserve them) but
'I would appreciate it greatly if you would offer some
'feedback, either on PSC or by email. Thank you. That way,
'I can make this a better programme, which would be great.

'This code is copyrighted and you may not resuse it in any
'way without my permission. Please respect that; I have
'been kind enough to share it with you in the first place.

'I recommend that you read this article through before you
'do anything else; just to get a feel for how the app
'works, if anything at all.

'Keep on coding! Oh, and keep it open source ;)

'Mischa Balen aka ~boltfish~


'              GLOBAL VARIABLES - Useage
'      ----------------------------------------

'OK, so firstly lets deal with the global variables,
'which are:

' ---------------
'| NumberOfTimes |
'| Binary        |
'| FileTemp      |
' ---------------

'1. NumberOfTimes =
'This is a global variable which the user can set from
'frmOptions.frm. It stores the data telling us how many
'times we should overwrite the file. Upon startup, its
'value is automatically 1000. If the user changes it,
'then it is updated and called by the ShredFile function.
'The maximum value is 999,999,999.

'2. Binary =

'The user can also specify whether or not to overwrite
'the files' contents with random binary data.
'It is accessed via frmOptions2.frm
'It is therefore a Boolean and is called by the ShredFile
'Routine.

'3. FileTemp =
'When the user clicks 'delete file' it goes into a loop
'until all the files have been removed by the ShredFile
'routine. Before the file is deleted, its path is written
'to FileTemp. Then, we use the GetFileName sub to return
'the file name. This is so we can add the file name to
'the status bar panel even if it has just been deleted.


'          APPLICATION MODEL AND FUNCTIONS:
'      ----------------------------------------

'Now that we have got those sorted out, let's take a look
'at what happens when the use clicks the delete button.
'Here is the basic model process for the whole app:

'Delete Clicked -> File Name Read -> File Encrypted ->
'File OverWritten -> File Replaced with "" -> File Deleted

'    DELETE CLICKED - We declare the following:
'    ------------------------------------------

    ' ------------------------
    '| Dim i As Integer       |
    '| Dim b As Integer       |
    '| Dim File2Del As String |
    '| Dim msg As String      |
    ' ------------------------

'1. I / B = Counter. B is labelled as the number of files
'in the listbox - i.e. the number of files to be deleted.
'I can be thought of as the current file (which is being
'deleted).

'Using I, we progress in steps of 1.

'In every stage, we:
'
'   1)Set the display panel to "Deleting" i "of" b
'   2)Set the other panel to the file name of the current
'     file which is being deleted
'   3)Use ShredFile Function

'This loop runs until I = B, i.e. when the current file
'being deleted = the total number of files. Therefore we
'must have finished, so we take appropriate action.


'     SHRED FILE FUNCTION - This is explained:
'     -----------------------------------------

'This is the primary and most important feature in the
'programme. It makes the file safe before finally deleting
'it. Again, look at the model below:

'   1. Generates random characters
'   2. Overwrites data in file with random characters
'   3. Does this until NumberOfTimes is satisfied
'------------END OF MAIN OVERWRITING LOOP------------
'   4. Checks to see if we should Binary the file (= True)
'   5. Corrupts the file IF Binary is True
'   6. Overwrites all the characters with ""
'   7. Deletes the file

'In order to make the random data, we generate a random
'number - Rnd(*255).
'Then we take this as a character code so we can convert it
'to a character data.

'Once this has been done, the file is opened for binary
'and we replace every character in it with the random
'character data we just generated.
'This process goes on until NumberOfTimes has been
'satisfied.
'We must flush the file buffers. If windoze sees that
'we are going to delete the file anyway it won't
'bother to overwrite it etc, so we use this API call
'in order to clear its "memory".

'If the user wants to hex corrupt the file (we look at
'the value of HexCorrupt), then we do so at this point.

'Finally, we remove all the data in the file and replace it
'with "".

'The file is now deleted.

'         BINARY FUNCTION - Explained below:
'     -----------------------------------------

'How this works:

'OK, it works by working in two loops:
'1. for x = 1 to NumberOfTimes
'2. for i = 1 to lof(1)

'In the second loop, it opens the file and overwrites
'each character with either a 0 or a 1 (specified by
'RandomBin(1) where 1 is the length). This is faster
'than the previous method of simply opening the file
'and generating a new string of length equal to the
'number of characters in the file.


'      ADDING CONTEXT MENUS - Explained below:
'     -----------------------------------------

'The File Shredder can add 'context menus' to files. The
'context menu is the menu that you get when you right
'click a file. A new option will be created called 'Delete
'with TFS' and when clicked, loads up the main programme
'with the file's path added to the main listbox.

'The user has the option to remove this function through
'the options menu. If the app path changes then they should
'choose to add the context menu from the options screen so
'that the registry is updated with the new value.


'      CUSTOM DELETION METHODS - Explained below:
'     -------------------------------------------

'From the options menu, the user can select one of five
'methods of deletion -

'1. Ultra Quick
'2. Quick
'3. Normal (default)
'4. Paranoid
'5. Custom

'When one of these is checked, the following global values
'are changed in accordance:

'Rename ----------- true / false - should we rename files?
'Binary ------- true / false - should we use binary on the files?
'Method ----------- what method? either shredfile or dod
'NumberOfTimes ---- number of times to overwrite the file using the hex corrupt
'Setting ---------- name e.g. normal/paranoid/custom

'When the user loads up the custom form to edit the details,
'the controls are filled in with the global values already
'specified by the current setting.













   
