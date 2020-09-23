<div align="center">

## Decent Beginner Tips


</div>

### Description

Just some handy little tips, vote if you want, thanks pscode.com, everyone here, and Nod Programming Inc.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[George E\.](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/george-e.md)
**Level**          |Beginner
**User Rating**    |4.2 (21 globes from 5 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/george-e-decent-beginner-tips__1-45014/archive/master.zip)





### Source Code


<b>Beware of optional typed parameters:</b>
Starting with Visual Basic 4.0, you could define optional parameters. There
was only one problem: They could be only of type Variant. With VB 5.0, you
can define typed optional parameters. However, you must be careful when doing
so, because you can't check whether a typed optional parameter was received.
Consider this sample code:
Public Sub SubX(Optional b As Boolean)
 If IsMissing(b) Then
  MsgBox "b is missing"
 Else
  MsgBox "b is not missing"
 End If
End Sub
...
 'Call SubX with no parameters
 SubX
You'd expect to see a message box indicating that b is missing, but no box
appears. The reason lies in the definition of IsMissing: "Returns a Boolean
value indicating whether an optional Variant argument has been passed to a
procedure." If you don't use a Variant argument, IsMissing won't provide the
expected value.
A typed optional parameter is never missing; it's always set to the default
value for each type (False for Boolean parameters, 0 for numbers and
zero-length strings).
Another option is to add the default value in the declaration of the
procedure, as follows:
Public Sub SubX(Optional i As Integer = 1)
****************************************************************************
<b>AVOIDING THE [ENTER] BEEP:</b>
When you're entering information into a text box and press [Enter], you'll
hear a beep. You can easily avoid this behavior. To do so, place a text box
on your form (Text1). Enter the following code in the KeyPress event:
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc(vbCr) Then
 KeyAscii = 0
End If
End Sub
When you run the form, pressing [Enter] will no longer produce a beep.
****************************************************************************
<b>Prevent partially painted windows:</b>
Sometimes when you display a form, only some of the controls appear. After a
pause, the remaining controls appear. Such partial painting doesn't look
professional. (Fortunately, this problem is much less apparent in VB 5.0
because of dramatic improvements in screen painting.)
To avoid partially painted windows when showing a non-modal form, use the
following code:
frmPerson.Show vbModeless
frmPerson.Refresh
The Refresh method will ensure that the form repainting is complete before
executing any other code in the routine.
****************************************************************************
<b>Centering a form:</b>
To center a form on the screen in VB3 or VB4, you can write a CenterForm
subroutine. Then, call CenterForm in the form's Load event. The code is as
follows:
Public Sub CenterForm(frmTarget As Form)
  frmTarget.Move (Screen.Width - frmTarget.Width) / 2, _
   (Screen.Height - frmTarget.Height) / 2
End Sub
Private Sub Form_Load()
  CenterForm Me
End Sub
Editor's Note:
In VB5, you can center a form on the screen by setting the StartUpPosition
property of the form to CenterScreen or CenterOwner.
****************************************************************************
<b>Case-conversion on the fly:</b>
If you want to convert text to uppercase as it's entered in a text box, just
create an Upper function and call it from the text box's keypress event, as
shown here:
  Private Sub Text1_KeyPress(KeyAscii As Integer)
  	KeyAscii = Upper(KeyAscii)
  End Sub
  Function Upper(KeyAscii As Integer)
  	If KeyAscii > 96 And KeyAscii < 123 Then
		KeyAscii = KeyAscii - 32
	End If
  	Upper = KeyAscii
  End Function
This technique eliminates the need to "UCase" entered data. It also makes
"hotseek" data searches much easier.
****************************************************************************
<b>Trapping dropdown list errors:</b>
In VB, the Text property of a Combo box whose Style property is set to
'2 - Dropdown List' is read-only. This means that a statement like:
MyCombo.Text = "The Third Item"
will return an error if "The Third Item" is not part of the list. Wouldn't
it be nice if VB just set the Combo box's ListIndex property to -1
(blanking it out) instead of bombing out? Well, here's some code that will
do just that:
Function SetComboText(MyCombo as ComboBox, MyItem as String) as Integer
 Dim I as Integer
 For I = 0 to MyCombo.ListCount - 1
 If MyCombo.List(I) = MyItem Then
  SetComboText = I
  Exit Function
 End If
 Next I
 ' If the program reaches this point, the string is not in the
 ' list.
 SetComboText = - 1
End Function
Use the function like this:
AnyCombo.ListIndex = SetComboText(AnyCombo, "Any String")
If "Any String" is in the list, then the combo box's ListIndex will be set
to the correct index; if not, it will be blanked out. The great thing about
this code is that if you want to do something else other than blanking out
the combo box, all you have to do is replace the line:
SetComboText = - 1
with whatever you wish.
****************************************************************************
<b>Speed up string buffers:</b>
Sometimes you need to write a program that builds up a large amount of data
in a string variable. You'd normally use a statement such as:
strBuffer = strBuffer & strNewData
during every loop. The problem with this approach is that the bigger your
string buffer becomes, the slower your program runs.
A neat and very simple way around this problem is to use another buffer.
Just fill the temporary buffer with data, and when it's big enough, append
it to the main buffer. Then, clear the temporary buffer and continue. The
code will look like this:
Public Sub NewBuildBuffer()
 Dim strBuffer As String, strTemp As String
 Dim l As Long, dStart As Date
 'Set start time
 dStart = Now
 'Build the buffer
 For l = 1 To 10000
  strTemp = strTemp & "New Line" & vbCrLf
  'Append to the main buffer every 100 times
  If l Mod 100 = 0 Then
   strBuffer = strBuffer & strTemp
   strTemp = ""
  End If
 Next
 'Append the last temp buffer
 strBuffer = strBuffer & strTemp
 'Report total time
 MsgBox "Seconds taken = " & DateDiff("s", dStart, Now)
End Sub
For programs that use very large string buffers, you'll see a huge
improvement.
****************************************************************************
<b>Preventing multiple instances of VB apps:</b>
You can easily prevent users from running multiple instances of your
programs by taking advantage of the PrevInstance property of the App object.
To do so, enter the following code in your application's opening form:
If App.PrevInstance Then
 MsgBox ("Cannot load program again."), vbExclamation, "The requested " _
  & "application is already open"
 Unload me
End If
This technique will also prevent multiple users from accessing single-user
applications.
****************************************************************************
<b>Retrieving the network logon name:</b>
You can easily retrieve a user's network logon name by using the following
API call:
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long
To retrieve a "clean" version of the name, use this function:
Public Function NTDomainUserName() As String
Dim strBuffer As String * 255
Dim lngBufferLength As Long
Dim lngRet As Long
Dim strTemp As String
	lngBufferLength = 255
	lngRet = GetUserName(strBuffer, lngBufferLength)
	strTemp = UCase(Trim$(strBuffer))
	NTDomainUserName = Left$(strTemp, Len(strTemp) - 1)
End Function
****************************************************************************
<b>Customizing a text box's pop-up menu:</b>
In Windows 95, right-clicking any text box brings up a context menu with
basic edit commands on it. If you want to change this menu, put the
following code in the MouseDown event of the text box.
If Button = vbRightButton Then
	Text1.Enabled = False
	Text1.Enabled = True
	Text1.SetFocus
	PopUpMenu Menu1
End If
where Text1 is the text box and Menu1 is the pop-up menu.
Disabling and re-enabling the control causes Windows to lose the MouseDown
message, SetFocus tidies things up a bit, and PopUpMenu shows the menu.
Left clicks will work as always, allowing the user to edit the text in the
text box.
****************************************************************************
<b>Selecting all text when a TextBox gets
focus:</b>
When you present the user with default text in a TextBox, you'll often want
to select that text when the TextBox gets focus. That way, the user can
easily type over your default text.
The function below will do the trick. The first click on the TextBox will
select all the text; the second click will place the cursor.
Public Sub TextSelected()
Dim i As Integer
Dim oMyTextBox As Object
Set oMyTextBox = Screen.ActiveControl
 If TypeName(oMyTextBox) = "TextBox" Then
  i = Len(oMyTextBox.Text)
  oMyTextBox.SelStart = 0
  oMyTextBox.SelLength = i
 End If
End Sub
Just add the function to your project and call it from the TextBox's
GotFocus event.
Private Sub Text1_GotFocus()
 TextSelected
End Sub
****************************************************************************
<b>Preventing Add-Ins from loading at launch:</b>
When you launch Visual Basic 4 or 5, any active Add-Ins also launch. If
there's an error in one of the Add-Ins, however, you could encounter a
global protection fault.
To prevent this from happening, you can turn off Add-Ins before launching
VB. To do so, launch Notepad or WordPad and open the file VBAddin.INI in
your Windows directory. You'll see a series of entries like this:
AppWizard.Wizard=1
Just change the "1" to a "0" in each entry. Then save the file and launch
VB. The program will launch without any Add-Ins.
Of course, to add and remove Add-Ins while you're in Visual Basic, just
choose Add-In Manager from the Add-Ins menu.
****************************************************************************
<b>Clearing all fields and combo boxes on a form:</b>
Sometimes you want to clear all the fields and combo boxes on a data-entry
form. If your form contains many controls, this could become tedious and
error prone. The following subroutine clears the contents of such fields on
your form automatically:
Public Sub ClearAllControls(frmForm As Form)
Dim ctlControl As Object
 ' Initialize all controls that can be initialized
 ' Any control with a text property or a list-index property
 On Error Resume Next
 For Each ctlControl In frmForm.Controls
  ctlControl.Text = ""
  ctlControl.ListIndex = -1
  DoEvents
 Next ctlControl
End Sub
Just call this procedure from your code like this:
Call ClearAllControls(Me)
****************************************************************************
<b>Quickly switching an object's Enabled property:</b>
You can easily switch an object's Enabled property with a single line of
code:
optSwitch.enabled = abs(optSwitch.enabled) - 1
Here's how the technique works: When Enabled is True, its numeric value is
-1. The absolute value of -1 is 1, so subtracting 1 from 1 would yield 0,
which is False. When Enabled is False, its numeric value is 0; 0 - 1
would then yield -1, or True.
This technique is an enhancement of the common usage
fraOption.enabled = optSwitch.enabled
to have an object follow the value of any other object's Enabled property.
Note: This technique depends on VB's definition of True and False.
To make this technique less dependent on that definition, you can use the
following code:
OptSwitch.enabled = NOT OptSwitch.enabled
This code works for any Boolean data type.
****************************************************************************
<b>Dealing with Null strings in Access database fields:</b>
By default Access string fields contain NULL values unless a string value
(including a blank string like "") has been assigned. When you read these
fields using recordsets into VB string variables, you get a runtime
type-mismatch error.
The best way to deal with this problem is to use the built-in & operator to
concatenate a blank string to each field as you read it. For example,
Dim DB As Database
Dim RS As Recordset
Dim sYear As String
Set DB = OpenDatabase("Biblio.mdb")
Set RS = DB.OpenRecordset("Authors")
sYear = "" & RS![Year Born]
****************************************************************************
<b>Specifying maximum lengths in a ComboBox:</b>
The ComboBox control doesn't have a MaxLength property like a TextBox does.
You can add some code to emulate this property, however. Just add the
following code to the KeyPress event of your ComboBox:
Private Sub Combo1_KeyPress(KeyAscii As Integer)
 'If the user is trying to type the eleventh key and...
 ' ...this key is not the Backspace Key, cancel the event!
 Const MAXLENGTH = 10
 If Len(Combo1.Text) >= MAXLENGTH And KeyAscii <> vbKeyBack Then
KeyAscii = 0
 '
End Sub
You can change the MaxLength value to any number you want. As you can see,
the code allows the user to use the [Backspace] key; you could enable other
keys by simply adding their KeyAscii values the way we did with [Backspace].
****************************************************************************
<b>Sharing resource files between VB and C projects:</b>
Suppose you want to use a resource file (RES) in your Visual Basic project,
but some of the file's resource indexes are greater than 0x8000. The VB
function LoadResString(index) receives an integer argument Index in the
range -32,768 to 32,767, so you can't pass values that are larger than
0x8000. You can solve this problem by passing the corresponding negative
index value, as follows (with 0 <=X < 0x8000):
RES   Visual Basic
0xFFFF - X  -X - 1=
0x8000+X  X-0x8000
For example, suppose you have the following RC file:
STRINGTABLE DISCARDABLE=
 BEGIN
 0xFFFF-0x0000 "resource string 1 with VB index -1 -0 = -1"
 0x8000+1  "resource string 2 with VB index - 32,768 + 1 = -32,767"
 END
To load string 1, you'll use LoadResString(-1). Similarly, to load string 2
you'll use LoadResString(-32767).
****************************************************************************
<b>The CDbl function versus Val:</b>
The Val() function is familiar, and it's useful for converting text box
numeric values to numbers. But if you use formatters to display large
numbers (with commas, for instance), there's a better function for your
purpose. The following examples illustrate the use of Val versus CDbl:
Code: print Val("12345")
Result: 12345
Code: print Val("12,345")
Result: 12
Code: print CDbl("12,345")
Result: 12345
Code: print CDbl("12345")
Result: 12345
Why are these functions different? The Visual Basic Help file offers
several hints. You should use the CDbl function instead of Val to provide
internationally aware conversions from any other data type to a Double. For
example, CDbl will recognize different decimal separators and thousands
separators properly depending on your system's locale.
Also, if you want your display and input routines to be automatically
reversible, you may want to consider using named numeric formats for
FORMAT(). Doing so helps guarantee a reversible process, given the LOCALE
setting of the user's machine.
****************************************************************************
<b>Command me, oh great one:</b>
Suppose you want to use Visual Basic to create an EXE that takes an input
value in a format like test.exe 2. Depending on the input value, you'll
perform certain tasks. In this situation, you can make use of the Command
function, which returns the argument portion of the command line you use to
launch VB or an EXE you develop in VB.
It's easy to send command-line information to an application. For instance,
to send information to an application called HappyApp, you could use the=
 line
HappyApp /CMD 1972
Now, within the application--probably in the Sub Main--you can use the
Command function to capture that command-line information.
To see this technique work, place a text box on a form. In the Form_Load
event, place the following line:
Text1.Text = Command
While still in VB, place some code on the command line. To do this in VB
3.0, choose Options | Project; in VB 4.0, choose Tools | Options..., then
click the Advanced tab; in VB 5.0, choose Project | Project Properties,
then click the Make tab. Next, type This is my argument in the Command Line
Arguments section and click OK. Run the application, and your command-line
text will appear in the text box.
Note that if you're working with 32-bit VB, I suggest creating an ActiveX
EXE or ActiveX DLL (formerly OLE Automation servers). By doing so, you
simply deal with property settings.
****************************************************************************
<b>Displaying and processing a message box</b>
The following code sample demonstrates an easy way to display and process a
message box (MsgBox) in any version of Visual Basic:
  Select Case MsgBox("Would you like to save the file somefile.txt?", _
  vbApplicationModal + vbQuestion + YesNoCancel, App.Title)
  Case vbYes
   'Save then file
  Case vbNo
   'Do something for No
  Case vbCancel
   'Do something else for Cancel
  End Select
This method works well, unless you need to save the answer from your Select
Case for later use. If you do, you'll need to use the more standard form of
prompting for the answer in a variable.
****************************************************************************
<b>Passing strings to a DLL:</b>
I recently came across a serious inefficiency in the way Visual Basic sends
strings to a DLL. The problem occurs when you want to get back a large
string field (32 KB) from a DLL written in C/C++. VB interacts somehow with
this string, causing significant overhead.
In order to call a DLL and get back a string-type data field, you must pass
a string and initialize it for as many bytes as you expect to be returned.
If you pass this function a small string, it will run quickly. But if you
pass it a large string (32 KB), the time will be significantly slower.
You'll see this slower performance even when no data is being returned,
meaning that the extra time results from some sort of VB overhead. As a
result, if speed is an issue when you're calling a DLL and passing a string
variable, you should pass a string that's only as large as you need.
You can find a sample project that demonstrates this problem in the file
Speed.zip at ftp.cobb.com/ivb/tipcode. The project simply loops for a
predetermined number of times and issues the standard windows API call
GetPrivateProfileString, which gets data from an INI file.
****************************************************************************
<b>Making a text box read-only:</b>
Here's a quick and easy way to make a text box read-only. Simply enter the
line
  keyascii = 0
in the textbox_keypress event.
The easiest way to make a text box read-only (in VB 4.0 and higher) is to
set the text box's Locked property to True. If you want to use our original
technique, you'll need to enter the code in the KeyDown event | KeyPress
doesn't trap the [Delete] key. However, if you don't set the Locked
property to True, Windows 95 will let you right-click on the text box to
open a context menu that gives you access to the Cut and Paste options.
****************************************************************************
<b>Creating a formless application:</b>
To create a VB program that has only console input and output--that is, no
dialog boxes or forms--you can use the Main procedure. Begin by creating a
new project. Open a code window, then choose Insert | Procedure.... In the
Insert Procedure dialog box, Select the Sub and Public options and enter
Main in the Name box. Click OK to create a new Main subroutine in the
General object. All your code will go in this routine; if you have any
useful BAS modules, you can add those to the project as well.
VB needs to know what code to execute when your application is called.
Since you're not using a form, you need to tell VB to start execution with
Sub Main. To do so, choose Tools | Options.... Click the Project tab and
select Sub Main from the Startup Form list. To remove the project's default
form, right-click on it in the Project window and choose Remove File from
the speed menu.
Testing a formless application can be a headache, so plan ahead: Use a log
file to get debug messages from your application. You'll want to read about
the Print # statement in VB's Help file, along with Open and Close.
Note that you can use this method to create a VB application that will run
as a service on NT. (Services can't have any forms or dialog boxes.)
****************************************************************************
<b>Case sensitivity in DLL calls:</b>
Use the Alias keyword to help convert non-case-sensitive VB 3.0 function
calls to their case-sensitive 32-bit counterparts.
When you declare or call a DLL in 32-bit Visual Basic, the name of the
function is case sensitive. To convert non-case-sensitive VB 3.0 calls to
case-sensitive calls, use the Alias keyword to hold the case-sensitive
function name. Place the name you want to call the function after the
Declare Sub/Function statement. (The Win32API.TXT file Aliases all function
calls, eliminating the case-sensitivity problem.)
****************************************************************************
<b>Simple input validation:</b>
Here's a way to achieve validation in text boxes and other controls that
support the KeyPress event. It's simple, but functional.
First, add this function to your project:
Function ValiText(KeyIn As Integer, _ValidateString As String, _Editable
 As Boolean) As Integer
 Dim ValidateList As String
 Dim KeyOut As Integer
 '
 If Editable = True Then
   ValidateList = UCase(ValidateString) & Chr(8)
 Else
   ValidateList = UCase(ValidateString)
 End If
 '
 If InStr(1, ValidateList, UCase(Chr(KeyIn)), 1) > 0 Then
  KeyOut = KeyIn
 Else
  KeyOut = 0
  Beep
 End If
 '
 ValiText = KeyOut
 '
End Function
Then, for each control whose input you wish to validate, just put something
like this in the KeyPress event of the control:
KeyAscii=ValiText(Keyascii, "0123456789/-",True)
Doing so will filter out any undesired keys that go to the control,
accepting only the keys defined by the second parameter. In this case, that
parameter ("0123456789/-") defines characters that are valid for a date.
The function's third parameter controls whether the [Backspace] key can be
used.
Note that this implementation of the function ignores the case of the
incoming keys, so if your second parameter were "abcdefg", the function
would also allow "ABCDEFG" to be entered.
****************************************************************************
<b>Simplying the addition of items to ComboBoxes:</b>
I often need to add items to a ComboBox and store an index or ID value in
the ItemData property. I've found that the code needed to add items to the
ComboBox and to check the ItemData property of the currently selected item
looks clumsy. So, I've written two simple helper routines to clean the code
up a bit. Here they are:
'---------------------------------------------------------------------------
 ' AddComboItem
 ' AddComboItem
'---------------------------------------------------------------------------
 Public Sub AddComboItem( _cboAdd As ComboBox, _ByVal sText As String,
 _ByVal lData As Long)
  cboAdd.AddItem sText
  cboAdd.ItemData(cboAdd.NewIndex) lData
 End Sub
'---------------------------------------------------------------------------
 ' CurrComboData
 ' CurrComboData
'---------------------------------------------------------------------------
 Public Function CurrComboData( _cbo As ComboBox) As Long
 If cbo.ListIndex <> -1 Then
  CurrComboData = cbo.ItemData(cbo.ListIndex)
 Else
  CurrComboData = -1
 End If
 End Function
Now, instead of writing
 cboTest.AddItem "Hello"
 cboTest.ItemData(cboTest.NewIndex) = 5
you can just write
 AddComboItem cboTest, "Hello",5
Instead of writing
 ID = cboTest.ItemData(cboTest.ListIndex)
you can write
 ID = CurrComboData( cboTest )
As an added bonus, CurrComboData protects you from the runtime error
generated if ListIndex is -1. Just be sure to check for a return of -1 from
CurrComboData.
****************************************************************************
<b>Showing long ListBox entries as a ToolTip:</b>
Sometimes the data you want to display in a list is too long for the size
of ListBox you can use. When this happens, you can use some simple code to
display the ListBox entries as ToolTips when the mouse passes over the
ListBox.
First, start a new VB project and add a ListBox to the default form. Then
declare the SendMessage API call and the constant (LB_ITEMFROMPOINT) needed
for the operation:
Option Explicit
'Declare the API function call.
Private Declare Function SendMessage _
 Lib "user32" Alias "SendMessageA" _
 (ByVal hwnd As Long, _
 ByVal wMsg As Long, _
 ByVal wParam As Long, _
 lParam As Any) As Long
' Add API constant
Private Const LB_ITEMFROMPOINT = &H1A9
Next, add some code to the form load event to fill the ListBox with data:
Private Sub Form_Load()
 '
 ' load some items in the list box
 With List1
  .AddItem "Michael Clifford Amundsen"
  .AddItem "Walter P.K. Smithworthy, III"
  .AddItem "Alicia May Sue McPherson-Pennington"
 End With
 '
End Sub
Finally, in the MouseMove event of the ListBox, put the following code:
Private Sub List1_MouseMove(Button As Integer, Shift As Integer, _
X As Single, Y As Single)
 '
 ' present related tip message
 '
 Dim lXPoint As Long
 Dim lYPoint As Long
 Dim lIndex As Long
 '
 If Button = 0 Then ' if no button was pressed
  lXPoint = CLng(X / Screen.TwipsPerPixelX)
  lYPoint = CLng(Y / Screen.TwipsPerPixelY)
  '
  With List1
   ' get selected item from list
   lIndex = SendMessage(.hwnd, _
    LB_ITEMFROMPOINT, _
    0, _
    ByVal ((lYPoint * 65536) + lXPoint))
   ' show tip or clear last one
   If (lIndex >= 0) And (lIndex <= .ListCount) Then
    .ToolTipText = .List(lIndex)
   Else
    .ToolTipText = ""
   End If
  End With '(List1)
 End If '(button=0)
 '
End Sub
****************************************************************************
<b>Creating Short Arrays Using the Variant Data Type:</b>
If you need to create a short list of items in an array, you can save a lot
of coding by using the Variant data type instead of a dimensioned standard
data type. This is especially handy when you need to create a list of short
phrases to support numeric output.
For example, add a button to a standard VB form and paste the following
code into the Click event of the button:
Private Sub Command1_Click()
 '
 ' create a quick array using variants
 '
 Dim aryList As Variant
 '
 aryList = Array("No Access", "Read-Only", "Update", "Delete")
 '
 MsgBox aryList(2)
 '
End Sub
****************************************************************************
<b>Using GetRows to Quickly Save Data Fields to Memory Variables:</b>
If you need to copy information from database fields into memory variables,
you can do it quickly using the GetRows method of the Recordset object. The
GetRows method copies one or more rows of data directly into a Variant data
type and stores the information as a two-dimensional array in the
formvarData(Field,Column).
To test the GetRow method, add a button to a VB form and paste the
following code into the Click event of the button. Be sure to fix the
reference to location of the BIBLIO.MDB database in the OpenDatabase
method. Also be sure to set up a reference to the Microsoft DAO 3.5 Object
Library.
Private Sub cmdGetDataRow_Click()
 '
 ' show getrow method
 '
 Dim ws As Workspace
 Dim db As Database
 Dim rs As Recordset
 '
 Dim varDataRows As Variant
 Dim intRows As Integer
 Dim intColumns As Integer
 '
 Dim intLoopRow As Integer
 Dim intLoopCol As Integer
 Dim strMsg As String
 '
 Set ws = DBEngine.CreateWorkspace(App.EXEName, "admin", "")
 Set db = ws.OpenDatabase("e:\devstudio\vb\biblio.mdb")
 Set rs = db.OpenRecordset("SELECT * FROM Authors")
 '
 intRows = InputBox("How Many Rows?", "GetRows Example", 0)
 intColumns = rs.Fields.Count
 varDataRows = rs.GetRows(intRows)
 '
 For intLoopRow = 0 To intRows - 1
  strMsg = ""
  For intLoopCol = 0 To intColumns - 1
   strMsg = strMsg & varDataRows(intLoopCol, intLoopRow) & vbCrLf
  Next
  MsgBox strMsg
 Next
 '
 rs.Close
 db.Close
 ws.Close
 '
End Sub
****************************************************************************
<b>Getting sensible Win32 API call errors:</b>
Most of the Win32 API calls return extended error information when they
fail. To get this information in a sensible format, you can use the
GetLastError and FormatMessage APIs.
Add the following declarations and function to a BAS module in a VB project:
Option Explicit
Public Declare Function GetLastError _
 Lib "kernel32" () As Long
Public Declare Function FormatMessage _
 Lib "kernel32" Alias "FormatMessageA" _
 (ByVal dwFlags As Long, _
 lpSource As Any, _
 ByVal dwMessageId As Long, _
 ByVal dwLanguageId As Long, _
 ByVal lpBuffer As String, _
 ByVal nSize As Long, _
 Arguments As Long) As Long
Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Public Function LastSystemError() As String
 '
 ' better system error
 '
 Dim sError As String * 500
 Dim lErrNum As Long
 Dim lErrMsg As Long
 '
 lErrNum = GetLastError
 lErrMsg = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, _
  ByVal 0&, lErrNum, 0, sError, Len(sError), 0)
 LastSystemError = Trim(sError)
 '
End Function
Now place a command button on a standard VB form and call the
LastSystemError function:
Private Sub Command1_Click()
 '
 MsgBox LastSystemError
 '
End Sub
If there was no error registered, you'll see a message saying "The
operation completed successfully."
When using this function, keep these points in mind:
1. Many API calls reset the value of GetLastError when successful, so the
function must be called immediately after the API call that failed.
2. The last error value is kept on a per-thread basis, therefore the
function must be called from the same thread as the API call that failed.
****************************************************************************
<b>Increment and decrement dates with the [+] and [-] keys:</b>
If you've ever used Quicken, you've probably notice a handy little feature
in that program's date fields. You can press the [+] key to increment one
day, [-] to decrement one day, [PgUp] to increment one month, and [PgDn] to
decrement one month. In this tip, we'll show you how to emulate this
behavior with Visual Basic.
First, insert a text box on a form (txtDate). Set its text property to ""
and its Locked property to TRUE.
Now place the following code in the KeyDown event:
Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
 '
 ' 107 = "+" KeyPad
 ' 109 = "-" KeyPad
 ' 187 = "+" (Actually this is the "=" key, same as "+" w/o the=
 shift)
 ' 189 = "-"
 ' 33 = PgUp
 ' 34 = PgDn
 '
 Dim strYear As String
 Dim strMonth As String
 Dim strDay As String
 '
 If txtDate.Text = "" Then
  txtDate.Text = Format(Now, "m/d/yyyy")
  Exit Sub
 End If
 '
 strYear = Format(txtDate.Text, "yyyy")
 strMonth = Format(txtDate.Text, "mm")
 strDay = Format(txtDate.Text, "dd")
 '
 Select Case KeyCode
  Case 107, 187 ' add a day
   txtDate.Text = Format(DateSerial(strYear, strMonth, strDay) +
1, "m/d/yyyy")
  Case 109, 189 ' subtract a day
   txtDate.Text = Format(DateSerial(strYear, strMonth, strDay) -
1, "m/d/yyyy")
  Case 33 ' add a month
   txtDate.Text = Format(DateSerial(strYear, strMonth + 1,
strDay), "m/d/yyyy")
  Case 34 ' subtract a month
   txtDate.Text = Format(DateSerial(strYear, strMonth - 1,
strDay), "m/d/yyyy")
 End Select
 '
End Sub
The one nasty thing about this is that if you have characters that are not
the characters usually in a date (i.e., 1-9, Monday, Tuesday, or /) you get
errors in the format command. To overcome this, I set the Locked property
to True. This way, the user can't actually type a character in the field,
but the KeyDown event still fires.
****************************************************************************
<b>Creating Win32 region windows:</b>
The Win32 API includes a really amazing feature called region windows. A
window under Win32 no longer has to be rectangular! In fact, it can be any
shape that may be constructed using Win32 region functions. Using the
SetWindowRgn Win32 function from within VB is so simple, but the results
are unbelievable! The following example shows a VB form that is NOT
rectangular!!
Here is the code. Enjoy!
 ' This goes into the General Declarations section:
Private Declare Function CreateEllipticRgn Lib "gdi32" _
 (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, _
 ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" _
 (ByVal hWnd As Long, ByVal hRgn As Long, _
 ByVal bRedraw As Boolean) As Long
Private Sub Form_Load()
Show 'The form!
SetWindowRgn hWnd, _
 CreateEllipticRgn(0, 0, 300, 200), _
 True
End Sub
****************************************************************************
<b>Manipulate your controls from the keyboard:</b>
If you're not comfortable using your mouse--or can't achieve the precise
results you'd like--these tips will come in handy.
First, you can resize controls at design time by using the [Shift] and
arrow keys, as follows:
  SHIFT + RIGHT ARROW increases the width of the control
  SHIFT + LEFT ARROW decreases the width of the control
  SHIFT + DOWN ARROW increases the height of the control
  SHIFT + UP ARROW decreases the height of the control
Note: The target control must have focus, so click on the control before
manipulating it from the keyboard.
Second, by using the [Control] key and the arrow keys, you can move your
controls at design time, as follows:
  CONTROL + RIGHT ARROW to move the control to the right
  CONTROL + LEFT ARROW to move the control to the left
  CONTROL + DOWN ARROW to move the control downwards
  CONTROL + UP ARROW to move the control upwards
If you select more than one control (by clicking on the first and
shift-clicking on the others), the above procedures will affect all the
selected controls.
****************************************************************************
<b>Simple file checking from anywhere:</b>
To keep my applications running smoothly, I often need to check that
certain files exist. So, I've written a simple routine to make sure they
do. Here it is:
Public Sub VerifyFile(FileName As String)
 '
 On Error Resume Next
 'Open a specified existing file
 Open FileName For Input As #1
 'Error handler generates error message with file and exits the routine
 If Err Then
  MsgBox ("The file " & FileName & " cannot be found.")
  Exit Sub
 End If
 Close #1
 '
End Sub
Now add a button to your form and place the code below behind the "Click"
event.
Private Sub cmdVerify_Click()
 '
 Call VerifyFile("MyFile.txt")
 '
End Sub
****************************************************************************
<b>Dragging items from one list to another:</b>
Here's a way that you can let users drag items from one list and drop them
in another one.
Create two lists (lstDraggedItems, lstDroppedItems) and a text box
(txtItem) in a form (frmTip).
Put the following code in the load event of your form.
Private Sub Form_Load()
 ' Set the visible property of txtItem to false
 txtItem.Visible = False
 'Add items to list1 (lstDraggedItems)
 lstDraggedItems.AddItem "Apple"
 lstDraggedItems.AddItem "Orange"
 lstDraggedItems.AddItem "Grape"
 lstDraggedItems.AddItem "Banana"
 lstDraggedItems.AddItem "Lemon"
 '
End Sub
In the mouseDown event of the list lstDraggedItems put the following code:
Private Sub lstDraggedItems_MouseDown(Button As Integer, Shift As Integer,
X As Single, Y As Single)
 '
 txtItem.Text = lstDraggedItems.Text
 txtItem.Top = Y + lstDraggedItems.Top
 txtItem.Left = X + lstDraggedItems.Left
 txtItem.Drag
 '
End Sub
In the dragDrop event of the list lstDroppedItems put the following code:
Private Sub lstDroppedItems_DragDrop(Source As Control, X As Single, Y As
Single)
 '
 If lstDraggedItems.ItemData(lstDraggedItems.ListIndex) = 9 Then
  Exit Sub
 End If
 ' To make sure that this item will not be selected again
 lstDraggedItems.ItemData(lstDraggedItems.ListIndex) = 9
 lstDroppedItems.AddItem txtItem.Text
 '
End Sub
Now you can drag items from lstDraggedItems and drop them in=
 LstDroppedItems.
Note that you cannot drag from the second list to the first. Also, the
dragged item remains in the first list. You'll have to address those
limitations yourself.
****************************************************************************
<b>Creating a new context menu in editable controls:</b>
This routine will permit you to replace the original context menu with your
private context menu in an editable control.
Add the following code to your form or to a BAS module:
Private Const WM_RBUTTONDOWN = &H204
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA"
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As
Any) As Long
Public Sub OpenContextMenu(FormName As Form, MenuName As Menu)
 'Tell system we did a right-click on the mdi
 Call SendMessage(FormName.hwnd, WM_RBUTTONDOWN, 0, 0&)
 'Show my context menu
 FormName.PopupMenu MenuName
 '
End Sub
Next, use the Visual Basic Menu Editor and the table below to create a
simple menu.
Caption		Name		Visible
Context Menu	mnuContext	NO
...First Item	mnuContext1
...Second Item	mnuContext2
Note that the last two items in the menu are indented (...) one level and
that only the first item in the list ("Context Menu") has the Visible
property set to NO.
Now add a text box to your form and enter the code below in the MouseDown
event of the text box.
Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As
Single, Y As Single)
 If Button = vbRightButton Then
  Call OpenContextMenu(Me, Me.mnuContext)
 End If
End Sub
Note: If you just want to kill the system context menu, just comment out
the line:
 FormName.PopupMenu MenuName
in the OpenContextMenu routine.
****************************************************************************
<b>Quick Custom Dialogs for DBGrid Cells:</b>
It's easy to add custom input dialogs to al the cells in the Microsoft Data
Bound Grid control.
First, add a DBGrid control and Data control to your form. Next, set the
DatabaseName and RecordSource properties of the data control to a valid
database and table ("biblio.mdb" and "Publishers" for example). Then set
the DataSource property of the DBGrid control to Data1 (the data control).
Now add the following code to your form.
' general declaration area
Dim strDBGridCell As String
Private Sub DBGrid1_AfterColEdit(ByVal ColIndex As Integer)
 '
 DBGrid1.Columns(ColIndex) = strDBGridCell
 '
End Sub
Private Sub DBGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii
As Integer, Cancel As Integer)
 '
 strDBGridCell = InputBox("Edit DBGrid Cell:", ,=
 DBGrid1.Columns(ColIndex))
 '
End Sub
Now whenever you attempt to edit any cell in the DBGrid, you'll see the
InputBox prompt you for input. You can replace the InputBox with any other
custom dialog you wish to build.
****************************************************************************
<b>Using the Alias Option to Prevent API Crashes:</b>
A number of Windows APIs have parameters that can be more than one data
type. For example, the WinHelp API call can accept the last parameter as a
Long or String data type depending on the service requested.
Visual Basic allows you to declare this data type as "Any" in the API call,
but this can lead to type mismatch errors or even system crashes if the
value is not the proper form.
You can prevent the errors and improve the run-time type checking by
declaring multiple versions of the same API function in your program. By
adding a function declaration for each possible parameter type, you can
continue to use strong data type checking.
To illustrate this technique, add the following APIs and constants to a
Visual Basic form. Notice that the two API declarations differ only in
their initial name ("WinHelp" and "WinHelpSearch") and the type declaration
of the last parameter ("dwData as Long" and "dwData as String").
' WinHelp APIs
Private Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd
As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData
As Long) As Long
Private Declare Function WinHelpSearch Lib "user32" Alias "WinHelpA" (ByVal
hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal
dwData As String) As Long
'
Private Const HELP_PARTIALKEY = &H105&
Private Const HELP_HELPONHELP = &H4
Private Const HelpFile = "c:\program files\devstudio\vb5\help\vb5.hlp"
Now add two command buttons to your form (cmdHelpAbout and cmdHelpSearch)
and place the following code behind the buttons. Be sure to edit the
location of the help file to match your installation of Visual Basic.
Private Sub cmdHelpAbout_Click()
 '
 WinHelp Me.hwnd, HelpFile, HELP_HELPONHELP, &H0
 '
End Sub
Private Sub cmdHelpSearch_Click()
 '
 WinHelpSearch Me.hwnd, HelpFile, HELP_PARTIALKEY, "option"
 '
End Sub
When you press on the HelpAbout button, you'll see help about using the
help system. When you press on the HelpSearch button, you'll see a list of
help entries on the "option" topic.
****************************************************************************
<b>Add Dithered Backgrounds to your VB Forms:</b>
Ever wonder how the SETUP.EXE screen gets its cool shaded background
coloring? This color shading is called dithering, and you can easily
incorporate it into your forms. Add the following routine to a form:
  Sub Dither(vForm As Form)
  Dim intLoop As Integer
   vForm.DrawStyle = vbInsideSolid
   vForm.DrawMode = vbCopyPen
   vForm.ScaleMode = vbPixels
   vForm.DrawWidth = 2
   vForm.ScaleHeight = 256
   For intLoop = 0 To 255
   vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 0,
255 -intLoop), B
   Next intLoop
  End Sub
Now, add to the Form_Activate event the line
  Dither ME
This version creates a fading blue background by adjusting the blue value
in the RGB function. (RGB stands for Red-Green-Blue.) You can create a
fading red background by changing the RGB call to
  RGB(255 - intLoop, 0, 0).
****************************************************************************
<b>Use FreeFile to Prevent File Open Conflicts:</b>
Both Access and VB let you hard code the file numbers when using the File
Open statement. For example:
  Open "myfile.txt" for Append as #1
  Print #1,"a line of text"
  Close #1
The problem with this method of coding is that you never know which file
numbers may be in use somewhere else in your program. If you attempt to use
a file number already occupied, you'll get a file error. To prevent this
problem, you should always use the FreeFile function. This function will
return the next available file number for your use. For example:
  IntFile=FreeFile()
  Open "myfile.txt" for Append as #intFile
  Print #intFile,"a line of text"
  Close #intFile
****************************************************************************
<b>Confirm Screen Resolution:</b>
Here's a great way to stop the user from running your application in the
wrong screen resolution. First, create a function called CheckRez:
Public Function CheckRez(pixelWidth As Long, pixelHeight As Long) As Boolean
 '
 Dim lngTwipsX As Long
 Dim lngTwipsY As Long
 '
 ' convert pixels to twips
 lngTwipsX = pixelWidth * 15
 lngTwipsY = pixelHeight * 15
 '
 ' check against current settings
 If lngTwipsX <> Screen.Width Then
  CheckRez = False
 Else
  If lngTwipsY <> Screen.Height Then
   CheckRez = False
  Else
   CheckRez = True
  End If
 End If
 '
End Function
Next, run the following code at the start of the program:
 If CheckRez(640, 480) = False Then
  MsgBox "Incorrect screen size!"
 Else
  MsgBox "Screen Resolution Matches!"
 End If
****************************************************************************
<b>Quick Text Select On GotFocus:</b>
When working with data entry controls, the current value in the control
often needs to be selected when the control received focus. This allows the
user to immediately begin typing over any previous value. Here's a quick
subroutine to do just that:
Public Sub FocusMe(ctlName As Control)
 With ctlName
  .SelStart = 0
  .SelLength = Len(ctlName)
 End With
End Sub
Now add a call to this subroutine in the GotFocus event of the input
 controls:
Private Sub txtFocusMe_GotFocus()
 Call FocusMe(txtFocusMe)
End Sub
****************************************************************************
<b>Use ParamArray to Accept an Arbitrary Number of Parameters:</b>
You can use the ParamArray keyword in the declaration line of a method to
create a subroutine or function that accepts an arbitrary number of
parameters at runtime. For example, you can create a method that will fill
a list box with some number of items even if you do not know the number of
items you will be sent. Add the method below to a form:
Public Sub FillList(ListControl As ListBox, ParamArray Items())
 '
 Dim i As Variant
 '
 With ListControl
  .Clear
  For Each i In Items
   .AddItem i
  Next
 End With
 '
End Sub
Note that the ParamArray keyword comes BEFORE the parameter in the
declaration line. Now add a list box to your form and a command button. Add
the code below in the "Click" event of the command button.
Private Sub Command1_Click()
 '
 FillList List1, "TiffanyT", "MikeS", "RochesterNY"
 '
End Sub
****************************************************************************
<b>Use FileDSNs to ease ODBC Installs:</b>
If you're using an ODBC connection to your database, you can ease the
process of installing the application on workstations by using the FileDSN
(data source name) instead of the more-common UserDSN. You define your ODBC
connection as you normally would with UserDSNs. However, the resulting
definition is not stored in the workstation registry. Instead it gets
stored in a text file with the name of the DSN followed by ".dsn" (i.e.
"MyFileDSN.dsn"). The default folder for all FileDSNs is "c:\program
files\common files\Odbc\data sources". Now, when you want to install the VB
application that uses the FileDSN, all you need to do is add the FileDSN to
the Install package and run the install as usual. No more setting up DSNs
manually!
NOTE: FileDSNs are available with ODBC 3.0 and higher.
****************************************************************************
<b>Opening a browser to your homepage</b>
You can use code like the following to open a browser to your homepage.
Modify filenames, paths, and URLs as necessary to match the values on your
system.
Dim FileName As String, Dummy As String
Dim BrowserExec As String * 255
Dim RetVal As Long
Dim FileNumber As Integer
Const SW_SHOWNORMAL =3D 1 ' Restores Window if Minimized or
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long
Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" _
(ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As _
String) As Long
'<Code> ---------
BrowserExec =3D Space(255)
FileName =3D "C:\temphtm.HTM"
FileNumber =3D FreeFile()  ' Get unused file number
Open FileName For Output As #FileNumber ' Create temp HTML file
 Write #FileNumber, "<HTML> <\HTML>" ' Output text
Close #FileNumber ' Close file
' Then find the application associated with it.
 RetVal =3D FindExecutable(FileName, Dummy, BrowserExec)
 BrowserExec =3D Trim$(BrowserExec)
 ' If an application is found, launch it!
 If RetVal <=3D 32 Or IsEmpty(BrowserExec) Then ' Error
 Msgbox "Could not find a browser"
 Else
 RetVal =3D ShellExecute(frmMain.hwnd, "open", BrowserExec, _
  "www.myurl.com", Dummy, SW_SHOWNORMAL)
 If RetVal <=3D 32 Then  ' Error
  Msgbox "Web Page not Opened"
 End If
 End If
Kill FileName ' delete temp HTML file
****************************************************************************
<b>Creating a incrementing number box</b>
You can't increment a vertical scroll bar's value--a fact that can become
annoying. For example, start a new project and place a text box and a
vertical scroll bar on the form. Place the vertical scroll bar to the right
of the text box and assign their Height and Top properties the same values.
Assign the vertical scroll bar a Min property value of 1 and a Max value of
10. Place the following code in the vertical scroll bar's Change event:
Text1.Text = VScroll1.Value
Now press [F5] to run the project. Notice that if you click on the bottom
arrow of the vertical scroll bar, the value increases; if you click on the
top arrow, the value decreases. From my perspective, it should be the other
way around.
To correct this, change the values of the Max and Min properties to
negative values. For example, end the program and return to the design
environment. Change the vertical scroll bar's Max value to -1 and its Min
value to -10. In its Change event, replace the line you entered earlier
with the following:
Text1.Text = Abs(Vscroll1.Value)
Now press [F5] to run the project. When you click on the top arrow of the
vertical scroll bar, the value now increases. Adjust the Height properties
of the text box and the scroll bar so you can't see the position indicator,
and your number box is ready to go.
****************************************************************************
<b>Measuring a text extent:</b>
It's very simple to determine the extent of a string in VB. You can do so
with WinAPI functions, but there's an easier way: Use the AutoSize property
of a Label component. First, insert a label on a form (labMeasure) and set
its AutoSize property to True and Visible property to False. Then write
this simple routine:
Private Function TextExtent(txt as String) as Integer
 labMeasure.Caption = txt
 TextExtent = labMeasure.Width
End Function
When you want to find out the extent of some text, simply call this
function with the string as a parameter.
In my case it turned out that the measure was too short. I just added some
blanks to the string. For example:
Private Function TextExtent(txt As String) As Integer
 labMeasure.Caption = " " & txt
 TextExtent = labMeasure.Width
End Function
****************************************************************************
<b>Importing Registry settings</b>
You can use just a few lines of code to import Registry settings. If you
have an application called myapp.exe and a Registry file called myapp.reg,
the following code will put those settings into the Registry without
bothering the user.
Dim strFile As String
strFile =3D App.Path & "\" & opts.AppExeName & ".reg"
If Len(Dir$(strFile)) > 1 Then
 lngRet =3D Shell("Regedit.exe /s " & strFile, vbNormalFocus)
End If
****************************************************************************
<b>Labeling your forms:</b>
Do you have a ton of screens in your application? Do you also have plenty
of users who want to "help you" by pointing out buttons that are one twip
out of place? Sometimes it's hard to know what screen users are talking
about when they're trying to communicate a problem--particularly if they're
in a different location than you.
To reduce the pain of this process, I add a label (called lblHeader) to the
top of each GUI window, nominally to hold start-up information for users
when they first open the window. You can also use this label to hold the
name of the window the user is looking at, by using the following code:
Private Sub Form_Load()
 SetupScreen me
End Sub
Public SetupScreen (frm as Form)
 ' Do other set-up stuff here (fonts, colors).
 HookInFormName frm
End Sub
Public Sub HookInFormName(frm As Form)
 ' The Resume Next on Error allows forms that do not use a standard
 ' header label to get past this.
 On Error Resume Next
 frm.lblHeader.Caption = "(" & frm.Name & ") " & frm.lblHeader.Caption
End Sub
Note that if you don't want to use a label, that you can also use code like
 frm.print frm.name
to print to the back of the window itself.
****************************************************************************

