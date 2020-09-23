VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmID3 
   Caption         =   "ID3"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6195
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   6195
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkv1 
      Caption         =   "Has ID3v1"
      Height          =   195
      Left            =   90
      TabIndex        =   33
      TabStop         =   0   'False
      ToolTipText     =   "Checked = contains ID3v1 info"
      Top             =   4440
      Width           =   1605
   End
   Begin VB.CheckBox chkv2 
      Caption         =   "Has ID3v2"
      Height          =   195
      Left            =   90
      TabIndex        =   32
      TabStop         =   0   'False
      ToolTipText     =   "Checked = contains ID3v2 info"
      Top             =   870
      Width           =   1605
   End
   Begin VB.Frame fraID3v1 
      Caption         =   "ID3v1 tag info"
      Height          =   3195
      Left            =   90
      TabIndex        =   24
      Top             =   4680
      Width           =   6045
      Begin VB.ComboBox cboGenre1 
         Height          =   360
         Left            =   3810
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1380
         Width           =   2145
      End
      Begin VB.CommandButton cmdUpdateID3v1tag 
         Caption         =   "Update"
         Enabled         =   0   'False
         Height          =   345
         Left            =   4830
         TabIndex        =   14
         ToolTipText     =   "Update the ID3v1 tag"
         Top             =   2550
         Width           =   1125
      End
      Begin VB.TextBox txtTrack1 
         Height          =   315
         Left            =   1770
         MaxLength       =   30
         TabIndex        =   8
         Top             =   210
         Width           =   4185
      End
      Begin VB.TextBox txtArtist1 
         Height          =   315
         Left            =   1770
         MaxLength       =   30
         TabIndex        =   9
         Top             =   600
         Width           =   4185
      End
      Begin VB.TextBox txtAlbum1 
         Height          =   315
         Left            =   1770
         MaxLength       =   30
         TabIndex        =   10
         Top             =   990
         Width           =   4185
      End
      Begin VB.TextBox txtGenre1 
         BackColor       =   &H8000000F&
         Height          =   360
         Left            =   1770
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1380
         Width           =   555
      End
      Begin VB.TextBox txtYear1 
         Height          =   315
         Left            =   1770
         MaxLength       =   4
         TabIndex        =   12
         Top             =   1770
         Width           =   4185
      End
      Begin VB.TextBox txtComments1 
         Height          =   315
         Left            =   1770
         MaxLength       =   30
         TabIndex        =   13
         Top             =   2160
         Width           =   4185
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "genre name:"
         Height          =   225
         Left            =   2520
         TabIndex        =   34
         Top             =   1440
         Width           =   1245
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         Caption         =   "song:"
         Height          =   225
         Left            =   240
         TabIndex        =   31
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         Caption         =   "artist:"
         Height          =   225
         Left            =   240
         TabIndex        =   30
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         Caption         =   "album:"
         Height          =   225
         Left            =   240
         TabIndex        =   29
         Top             =   1050
         Width           =   1455
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "genre #:"
         Height          =   225
         Left            =   900
         TabIndex        =   28
         Top             =   1440
         Width           =   795
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         Caption         =   "year:"
         Height          =   225
         Left            =   240
         TabIndex        =   27
         Top             =   1830
         Width           =   1455
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "comments:"
         Height          =   225
         Left            =   240
         TabIndex        =   26
         Top             =   2220
         Width           =   1455
      End
   End
   Begin VB.Frame fraID3v2 
      Caption         =   "ID3v2 tag info"
      Height          =   3195
      Left            =   90
      TabIndex        =   15
      Top             =   1140
      Width           =   6045
      Begin VB.CommandButton cmdCopyToV1 
         Caption         =   "Copy to ID3v1"
         Enabled         =   0   'False
         Height          =   345
         Left            =   4500
         TabIndex        =   36
         ToolTipText     =   "Update the ID3v1 tag"
         Top             =   2640
         Width           =   1455
      End
      Begin VB.ComboBox cboGenre2 
         Height          =   360
         Left            =   3810
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1470
         Width           =   2145
      End
      Begin VB.TextBox txtTrackNbr 
         Height          =   315
         Left            =   1770
         TabIndex        =   7
         Top             =   2640
         Width           =   585
      End
      Begin VB.TextBox txtTrack 
         Height          =   315
         Left            =   1770
         TabIndex        =   1
         Top             =   330
         Width           =   4185
      End
      Begin VB.TextBox txtArtist 
         Height          =   315
         Left            =   1770
         TabIndex        =   2
         Top             =   715
         Width           =   4185
      End
      Begin VB.TextBox txtAlbum 
         Height          =   315
         Left            =   1770
         TabIndex        =   3
         Top             =   1100
         Width           =   4185
      End
      Begin VB.TextBox txtGenre 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1485
         Width           =   555
      End
      Begin VB.TextBox txtYear 
         Height          =   315
         Left            =   1770
         TabIndex        =   5
         Top             =   1870
         Width           =   4185
      End
      Begin VB.TextBox txtComments 
         Height          =   315
         Left            =   1770
         TabIndex        =   6
         Top             =   2255
         Width           =   4185
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "genre name:"
         Height          =   225
         Left            =   2430
         TabIndex        =   35
         Top             =   1530
         Width           =   1245
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "track nbr:"
         Height          =   225
         Left            =   840
         TabIndex        =   23
         Top             =   2670
         Width           =   885
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "song:"
         Height          =   225
         Left            =   270
         TabIndex        =   22
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "artist:"
         Height          =   225
         Left            =   270
         TabIndex        =   21
         Top             =   750
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "album:"
         Height          =   225
         Left            =   270
         TabIndex        =   20
         Top             =   1140
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "genre #:"
         Height          =   225
         Left            =   930
         TabIndex        =   19
         Top             =   1530
         Width           =   795
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "year:"
         Height          =   225
         Left            =   270
         TabIndex        =   18
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "comments:"
         Height          =   225
         Left            =   270
         TabIndex        =   17
         Top             =   2310
         Width           =   1455
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5580
      Top             =   510
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Get ID3 info for MP3 file..."
      Height          =   495
      Left            =   90
      TabIndex        =   0
      ToolTipText     =   "Select an MP3 you want to read."
      Top             =   30
      Width           =   2595
   End
End
Attribute VB_Name = "frmID3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Last Modified: April 27, 2001
' Contact: Kevin Pohl
' email: kevinpohl@threefifteen.net
'
' Code Purpose:
' 1. To read ID3v1 tag info
' 2. To read ID3v2 tag info
' 3. To update ID3v1 tags
'
' Basically, I'm trying to extract this info from ID3v1 and
' ID3v2 tags:
' 1. Track Title
' 2. Artist
' 3. Album
' 4. Genre number
' 5. Genre name (use the genre number to get the genre name)
' 6. Track number (so far, only for ID3v2)
' 7. Comments
' 8. Year

' Missing features:
' 1. Doesn't update ID3v2 tags
'
' Features to be added:
' 1. Writing ID3v2 tags
' 2. This code currently works after selecting a file with a
'    common dialog.  I'm going to expand on this code to
'    batch read MP3 files after selecting MP3s in a list box
'    that is populated from a recursive directory search.
'    Obviously, this code can serve as the foundation for that
'    but will have to be modified.
'
' Wish list:
' 1. To be able to write/update more than the basic tag info
'    for ID3v2 tags, including image, lyrics, MusicMatch fields
'    (Tempo, Situation, Preference, Mood), and URL link fields.
' 2. To differentiate the difference ID3v1 tags that contain a
'    track number and those that do not.
'
'
' Known issues:
' 1. This code isn't reading the 'Comment' field of ID3v2
'    tags properly.
' 2. This code isn't reading ID3 tags of MP3 files that
'    were created by RealJukebox.  RJB does something
'    different that I haven't totally figured out yet.
'    RJB seems to fill the tag with a bunch of info
'    in the 'GEOB" frame of the tag, which makes it
'    difficult to extract the info.  Most of the software
'    I've used that reads tags doesn't read RJB MP3 file
'    tags.
' 3. Doesn't read or write ID3v2 tags from Windows Media
'    Audio files.
' 4. Some MP3-related software (like MusicMatch) tags
'    the ID3v2 genre as "(17) General Rock".  I've written
'    some code that parses the ID3v2 genre to determine if
'    parentheses exists in the genre "()" and to extract
'    only the genre number.  I haven't plugged this feature
'    into this code yet.
''
' Other comments:
' 1. I'm using VB 6.0 Enterprise Edition w/SP4
' 2. This code should work with any version of VB with
'    any SP.
' 3. A compile exe of this code should run on all
'    operating systems (Win95/98/Me/NT/2000)
'
' Code sources:
' 1. Code for reading and writing ID3v1 tags is basically
'    all the same.  You can find variations on:
'    http://www.planet-source-code
'    http://www.freevbcode.com
' 2. I found code written by a guy with the alias of
'    "The Frog Prince".  He has written VB code that utilizes
'    a DLL named VBID3LIB.dll.  It can read and write ID3v2
'    tags.  But it's main weakness is that it will read the
'    ID3v1 tag instead of the ID3v2 tag if the MP3 contains
'    both ID3v2 AND an ID3v2 tag.  You can find the code at
'    his web site:
'    http://members.tripod.com/thefrogprince/id3lib.htm
' 3. There's also some code contained in a VB project
'    written by "Joe Hart".  His code can be found on various
'    VB source code sites, including http://www.freevbcode.com
'    This is the link to his project at freevbcode:
'    http://www.freevbcode.com/ShowCode.Asp?ID=1127
'    This code reads ID3v1 and ID3v2 tags, and supposedly
'    writes ID3v1 tags.  His project is an MP3 player and file
'    manager.  But I found a lot of problems with it and very
'    difficult to follow.  I've extracted what I believe is the
'    "good stuff" in his code to read tags and have include it
'    in this project.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public i As Integer
Public strEmptyString As String
Public B As Byte
Public s As String

Private Type ID3v1Tag
  id As String * 3
  title As String * 30
  Artist As String * 30
  Album As String * 30
  Year As String * 4
  Comment As String * 30
  Genre As Byte
End Type

'Private Tag1 As ID3v1Tag

Private Version As Byte
Private Const sGenreMatrix = " |Blues|Classic Rock|Country|Dance|Disco|Funk|Grunge|" + _
    "Hip-Hop|Jazz|Metal|New Age|Oldies|Other|Pop|R&B|Rap|Reggae|Rock|Techno|" + _
    "Industrial|Alternative|Ska|Death Metal|Pranks|Soundtrack|Euro-Techno|" + _
    "Ambient|Trip Hop|Vocal|Jazz+Funk|Fusion|Trance|Classical|Instrumental|Acid|" + _
    "House|Game|Sound Clip|Gospel|Noise|Alt. Rock|Bass|Soul|Punk|Space|Meditative|" + _
    "Instrumental Pop|Instrumental Rock|Ethnic|Gothic|Darkwave|Techno-Industrial|Electronic|" + _
    "Pop-Folk|Eurodance|Dream|Southern Rock|Comedy|Cult|Gangsta Rap|Top 40|Christian Rap|" + _
    "Pop/Punk|Jungle|Native American|Cabaret|New Wave|Phychedelic|Rave|Showtunes|Trailer|" + _
    "Lo-Fi|Tribal|Acid Punk|Acid Jazz|Polka|Retro|Musical|Rock & Roll|Hard Rock|Folk|" + _
    "Folk/Rock|National Folk|Swing|Fast-Fusion|Bebob|Latin|Revival|Celtic|Blue Grass|" + _
    "Avantegarde|Gothic Rock|Progressive Rock|Psychedelic Rock|Symphonic Rock|Slow Rock|" + _
    "Big Band|Chorus|Easy Listening|Acoustic|Humour|Speech|Chanson|Opera|Chamber Music|" + _
    "Sonata|Symphony|Booty Bass|Primus|Porn Groove|Satire|Slow Jam|Club|Tango|Samba|Folklore|" + _
    "Ballad|power Ballad|Rhythmic Soul|Freestyle|Duet|Punk Rock|Drum Solo|A Cappella|Euro-House|" + _
    "Dance Hall|Goa|Drum & Bass|Club-House|Hardcore|Terror|indie|Brit Pop|Negerpunk|Polsk Punk|" + _
    "Beat|Christian Gangsta Rap|Heavy Metal|Black Metal|Crossover|Comteporary Christian|" + _
    "Christian Rock|Merengue|Salsa|Trash Metal|Anime|JPop|Synth Pop"
Private GenreArray() As String


Private Sub cmdCopyToV1_Click()

txtTrack1.Text = Trim(Left$(txtTrack.Text, 30))
txtArtist1.Text = Trim(Left$(txtArtist.Text, 30))
txtAlbum1.Text = Trim(Left$(txtAlbum.Text, 30))
txtYear1.Text = Trim(Left$(txtYear.Text, 4))
txtComments1.Text = Trim(Left$(txtComments.Text, 30))
cboGenre1.ListIndex = cboGenre2.ListIndex
cmdUpdateID3v1tag.Enabled = True

End Sub

Private Sub Form_Load()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Populate the ID3v1 and ID3v2 genre combo boxes
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Populate the genre array
        GenreArray = Split(sGenreMatrix, "|")
        
    'Populate the combo boxes
        For i = LBound(GenreArray) To UBound(GenreArray)
            cboGenre1.AddItem GenreArray(i)
            cboGenre2.AddItem GenreArray(i)
        Next

End Sub
Private Sub cmdBrowse_Click()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Open the CommonDialog to choose an MP3 file
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim strFilePath As String
    strFilePath = FileChoice(0, "Choose MP3 file")
    If strFilePath <> "" Then
        Me.Caption = strFilePath
    End If
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Clear all fields on the window
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'ID3v1 text fields
        chkv1.Value = 0         'This becomes checked if the file has an ID3v1 tag
        txtTrack1.Text = ""     'The file's track title (limited to 30 characters)
        txtArtist1.Text = ""    'The file's artist (limited to 30 characters)
        txtAlbum1.Text = ""     'The file's album (limited to 30 characters)
        txtGenre1.Text = ""     'The file's genre
        cboGenre1.ListIndex = 0
        txtYear1.Text = ""      'The file's year (limited to 4 characters)
        txtComments1.Text = ""  'The file's comments (limited to 30 characters)
        cmdUpdateID3v1tag.Enabled = False
        
    'ID3v2 text fields
        chkv2.Value = 0         'This becomes checked if the file has an ID3v2 tag
        txtTrack.Text = ""      'The file's track title (not limited to 30 characters)
        txtArtist.Text = ""     'The file's artist (not limited to 30 characters)
        txtAlbum.Text = ""      'The file's album (not limited to 30 characters)
        txtGenre.Text = ""      'The file's genre (not limited to 30 characters)
        cboGenre2.ListIndex = 0
        txtYear.Text = ""       'The file's year (not limited to 4 characters)
        txtComments.Text = ""   'The file's comments (not limited to 30 characters)
        txtTrackNbr.Text = ""   'The file's track number (not limited to 30 characters)
        cmdCopyToV1.Enabled = False
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Read the file to extract any existing ID3v1 and ID3v2 tags
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ReadFile

End Sub

Private Function ReadFile()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This function:
' 1) Opens the file that was selected in the common dialog.
' 2) Checks for a valid ID3 header.
' 3) Extracts any ID3v1 tag info (and displays it on window)
' 4) Extracts anu ID3v2 tag info (and displays it on window)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

On Error GoTo errorhandler

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' use the filename to get ID3 info
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim strFilename As String
    Dim lngFilesize As Long
    
    strFilename = Me.Caption

    Dim fn As Integer
    Dim lngHeaderPosition As Long
    Dim Tag1 As ID3v1Tag
    Dim Tag2 As String
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Open the file
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    fn = FreeFile
    
    Open strFilename For Binary As #fn                      'Open the file so we can read it
    lngFilesize = LOF(fn)                                   'Size of the file, in bytes

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Check for a Header
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
        Get #fn, 1, B
            
        If B <> 255 Then '(255 is where an ID3v2 header should start)
            If B <> 73 Then
                'Exit Function
            End If
        End If
         
        lngHeaderPosition = 1
        Get #fn, 2, B
        If (B < 250 Or B > 251) Then
            'We have an ID3v2 tag
            cmdCopyToV1.Enabled = True
            If B = 68 Then
                Get #fn, 3, B
                If B = 51 Then
                    Dim R As Double
                    Get #fn, 4, Version
                    Get #fn, 7, B
                    R = B * 20917152
                    Get #fn, 8, B
                    R = R + (B * 16384)
                    Get #fn, 9, B
                    R = R + (B * 128)
                    Get #fn, 10, B
                    R = R + B
                    If R > lngFilesize Or R > 2147483647 Then
                        Exit Function
                    End If
                    Tag2 = Space$(R)
                    Get #fn, 11, Tag2
                    lngHeaderPosition = R + 11
                End If
            End If
        Else
            'ID3v2 tag is missing
        End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Check for an ID3v1 tag
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'ID3v1 tag
        Get #fn, lngFilesize - 127, Tag1.id
        
        If Tag1.id = "TAG" Then 'If "TAG" is present, then we have a valid ID3v1 tag and will extract all available ID3v1 info from the file
            cmdUpdateID3v1tag.Enabled = True
            Get #fn, , Tag1.title   'Always limited to 30 characters
            Get #fn, , Tag1.Artist  'Always limited to 30 characters
            Get #fn, , Tag1.Album   'Always limited to 30 characters
            Get #fn, , Tag1.Year    'Always limited to 4 characters
            Get #fn, , Tag1.Comment 'Always limited to 30 characters
            Get #fn, , Tag1.Genre   'Always limited to 1 byte (?)
            
            frmID3.chkv1.Value = 1 'Indicates that the file contains ID3v1 tag info
    
            'Populate the form with the ID3v1 info
            With frmID3
                txtTrack1.Text = Trim$(Tag1.title)
                txtArtist1.Text = Trim$(Tag1.Artist)
                txtAlbum1.Text = Trim$(Tag1.Album)
                txtYear1.Text = Trim$(Tag1.Year)
                txtComments1.Text = Trim$(Tag1.Comment)
                getgenrefromID (Tag1.Genre)
            End With

        End If
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Proceed to extract the ID3v2 tag info if any exists
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     
        If Tag2 <> strEmptyString Then
            frmID3.chkv2.Value = 1
            GetID3v2Tag1 (Tag2) 'Pass the Id3v2 TagId to the GetID3v2Tag1 function
        End If
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Close the file
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Close

    Exit Function
        
errorhandler:
    'MsgBox "Error reading file"
    Err.Clear
    Close
    Resume Next
End Function

Private Function GetID3v2Tag1(Tag2 As String) As Boolean

On Error GoTo errorhandler

   Dim TitleField As String
   Dim ArtistField As String
   Dim AlbumField As String
   Dim YearField As String
   Dim GenreField As String
   Dim FieldSize As Long
   Dim SizeOffset As Long
   Dim FieldOffset As Long
   Dim TrackNbr As String
   Dim SituationField As String
   
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Determine if the ID3v2 tag is ID3v2.2 or ID3v2.3
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Notes: I haven't tested reading an MP3 file that has a ID3v2.2 tag
   
    Select Case Version
    
        Case 2 'ID3v2.2
        'Set the fieldnames for version 2.0
            TitleField = "TT2"
            ArtistField = "TOA"
            AlbumField = "TAL"
            YearField = "TYE"
            GenreField = "TCO"
            FieldOffset = 7
            SizeOffset = 5
            TrackNbr = "TRCK"
       
        Case 3 'ID3v2.3
        'Set the fieldnames for version 3.0
            TitleField = "TIT2"
            ArtistField = "TPE1"
            AlbumField = "TALB"
            YearField = "TYER"
            GenreField = "TCON"
            TrackNbr = "TRCK"
       
            FieldOffset = 11
            SizeOffset = 7
        Case Else
        'We don't have a valid ID3v2 tag, so bail
            Exit Function
            
    End Select
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Extract track title
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
       i = InStr(Tag2, TitleField)
       If i > 0 Then
          'read the title
          FieldSize = Asc(Mid$(Tag2, i + SizeOffset)) - 1
          If Version = 3 Then
             'check for compressed or encrypted field
             B = Asc(Mid$(Tag2, i + 9))
             If (B And 128) = True Or (B And 64) = True Then GoTo ReadAlbum
          End If
          txtTrack.Text = Mid$(Tag2, i + FieldOffset, FieldSize)
       End If
       
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Extract album title
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ReadAlbum:
    i = InStr(Tag2, AlbumField)
    If i > 0 Then
       FieldSize = Asc(Mid$(Tag2, i + SizeOffset)) - 1
       If Version = 3 Then
          'check for compressed or encrypted field
          B = Asc(Mid$(Tag2, i + 9))
          If (B And 128) = 128 Or (B And 64) = 64 Then GoTo ReadArtist
       End If
       txtAlbum.Text = Mid$(Tag2, i + FieldOffset, FieldSize)
       
    End If
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Extract artist name
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ReadArtist:
   i = InStr(Tag2, ArtistField)
   If i > 0 Then
      FieldSize = Asc(Mid$(Tag2, i + SizeOffset)) - 1
      If Version = 3 Then
         'check for compressed or encrypted field
         B = Asc(Mid$(Tag2, i + 9))
         If (B And 128) = 128 Or (B And 64) = 64 Then GoTo ReadYear
      End If
      txtArtist.Text = Mid$(Tag2, i + FieldOffset, FieldSize)
   End If
   
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Extract year title
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ReadYear:
   i = InStr(Tag2, YearField)
   If i > 0 Then
      FieldSize = Asc(Mid$(Tag2, i + SizeOffset)) - 1
      If Version = 3 Then
         'check for compressed or encrypted field
         B = Asc(Mid$(Tag2, i + 9))
         If (B And 128) = 128 Or (B And 64) = 64 Then GoTo ReadGenre
      End If
      txtYear.Text = Mid$(Tag2, i + FieldOffset, FieldSize)
   End If
   
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Extract genre
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ReadGenre:
   i = InStr(Tag2, GenreField)
   If i > 0 Then
      FieldSize = Asc(Mid$(Tag2, i + SizeOffset)) - 1
      
      If Version = 3 Then
         'check for compressed or encrypted field
         B = Asc(Mid$(Tag2, i + 9))
         If (B And 128) = 128 Or (B And 64) = 64 Then GoTo ReadTrackNbr
      End If
      
      s = Mid$(Tag2, i + FieldOffset, FieldSize)
      
      If Left$(s, 1) = "(" Then
      
        'we have an ID3v2 genre that's in the format of:
            '(0) Blues
            '(12) Rock
            '(125) Dance Hall
            
            Dim intStrip As Integer
            Dim intParPos As Integer
            Dim strStrip As String
            
            intParPos = InStr(1, s, ")", 0)
            strStrip = Trim(Right(s, (Len(s) - intParPos)))
            cboGenre2.Text = strStrip
        
      Else
         
         If i > 0 Then
            txtGenre.Text = "n/a"
            cboGenre2.Text = s
         End If
      End If
      
   End If
   
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Extract track number
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ReadTrackNbr:
   i = InStr(Tag2, TrackNbr)
   If i > 0 Then
      FieldSize = Asc(Mid$(Tag2, i + SizeOffset)) - 1
      If Version = 3 Then
         'check for compressed or encrypted field
         B = Asc(Mid$(Tag2, i + 9))
         If (B And 128) = 128 Or (B And 64) = 64 Then GoTo Done
      End If
      txtTrackNbr.Text = Mid$(Tag2, i + FieldOffset, FieldSize)
   End If
   
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' We're done looking for ID3v2 info
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Done:
   
   Exit Function

errorhandler:
   Err.Clear
   Resume Next
End Function

Private Function FileChoice(iFileType As Integer, title As String) As String

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Prepares the common dialog
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo errorhandler

    With CommonDialog1
        .DialogTitle = "Choose an MP3 audio file"           'Sets the caption of the ShowOpen common dialog
        .CancelError = True                                 'Sets a value indicating whether an error is generated when the user chooses the Cancel button.
        .Filter = "MP3 (*.mp3)|*.mp3|"                      'Sets the filters that are displayed in the Typelist box of a dialog box.
        .DefaultExt = ".mp3"                                'Sets the default filename extension for the dialog box.
        .FilterIndex = 2                                    'Sets the default filter for the dialog box.
        .ShowOpen                                           'Displays the CommonDialog control's Open dialog box.
    End With
    
    FileChoice = CommonDialog1.FileName                     'Returns the path and filename of a selected file
    Exit Function
    
errorhandler:
    'MsgBox "Error selecting MP3 audio file."
    Err.Clear
    Resume Next
    
End Function

Private Sub cmdUpdateID3v1tag_Click()

Dim strFilename As String
strFilename = Me.Caption

Dim Tag As String * 3
Dim Songname As String * 30
Dim Artist As String * 30
Dim Album As String * 30
Dim Year As String * 4
Dim Comment As String * 30
Dim Genre As Byte

Tag = "TAG"
Songname = txtTrack1.Text
Artist = txtArtist1.Text
Album = txtAlbum1.Text
Year = txtYear1.Text
Comment = txtComments1.Text

'Get ID from genre name selected in listbox
    Dim strListGenre As String
    strListGenre = cboGenre1.Text
    
    If strListGenre = "Blues" Then
        Genre = 0
    ElseIf strListGenre = "Classic Rock" Then
        Genre = 1
    ElseIf strListGenre = "Country" Then
        Genre = 2
    ElseIf strListGenre = "Dance" Then
        Genre = 3
    ElseIf strListGenre = "Disco" Then
        Genre = 4
    ElseIf strListGenre = "Funk" Then
        Genre = 5
    ElseIf strListGenre = "Grunge" Then
        Genre = 6
    ElseIf strListGenre = "Hip-Hop" Then
        Genre = 7
    ElseIf strListGenre = "Jazz" Then
        Genre = 8
    ElseIf strListGenre = "Metal" Then
        Genre = 9
    ElseIf strListGenre = "New Age" Then
        Genre = 10
    ElseIf strListGenre = "Oldies" Then
        Genre = 11
    ElseIf strListGenre = "Other" Then
        Genre = 12
    ElseIf strListGenre = "Pop" Then
        Genre = 13
    ElseIf strListGenre = "R&B" Then
        Genre = 14
    ElseIf strListGenre = "Rap" Then
        Genre = 15
    ElseIf strListGenre = "Reggae" Then
        Genre = 16
    ElseIf strListGenre = "Rock" Then
        Genre = 17
    ElseIf strListGenre = "Techno" Then
        Genre = 18
    ElseIf strListGenre = "Industrial" Then
        Genre = 19
    ElseIf strListGenre = "Alternative" Then
        Genre = 20
    ElseIf strListGenre = "Ska" Then
        Genre = 21
    ElseIf strListGenre = "Death Metal" Then
        Genre = 22
    ElseIf strListGenre = "Pranks" Then
        Genre = 23
    ElseIf strListGenre = "Soundtrack" Then
        Genre = 24
    ElseIf strListGenre = "Euro-Techno" Then
        Genre = 25
    ElseIf strListGenre = "Ambient" Then
        Genre = 26
    ElseIf strListGenre = "Trip-Hop" Then
        Genre = 27
    ElseIf strListGenre = "Vocal" Then
        Genre = 28
    ElseIf strListGenre = "Jazz+Funk" Then
        Genre = 29
    ElseIf strListGenre = "Fusion" Then
        Genre = 30
    ElseIf strListGenre = "Trance" Then
        Genre = 31
    ElseIf strListGenre = "Classical" Then
        Genre = 32
    ElseIf strListGenre = "Instrumental" Then
        Genre = 33
    ElseIf strListGenre = "Acid" Then
        Genre = 34
    ElseIf strListGenre = "House" Then
        Genre = 35
    ElseIf strListGenre = "Game" Then
        Genre = 36
    ElseIf strListGenre = "Sound Clip" Then
        Genre = 37
    ElseIf strListGenre = "Gospel" Then
        Genre = 38
    ElseIf strListGenre = "Noise" Then
        Genre = "39"
    ElseIf strListGenre = "AlternRock" Then
        Genre = 40
    ElseIf strListGenre = "Bass" Then
        Genre = 41
    ElseIf strListGenre = "Soul" Then
        Genre = 42
    ElseIf strListGenre = "Punk" Then
        Genre = 43
    ElseIf strListGenre = "Space" Then
        Genre = 44
    ElseIf strListGenre = "Meditative" Then
        Genre = 45
    ElseIf strListGenre = "Instrumental Pop" Then
        Genre = 46
    ElseIf strListGenre = "Instrumental Rock" Then
        Genre = 47
    ElseIf strListGenre = "Ethnic" Then
        Genre = 48
    ElseIf strListGenre = "Gothic" Then
        Genre = 49
    ElseIf strListGenre = "Darkwave" Then
        Genre = 50
    ElseIf strListGenre = "Techno-Industrial" Then
        Genre = 51
    ElseIf strListGenre = "Electronic" Then
        Genre = 52
    ElseIf strListGenre = "Pop-Folk" Then
        Genre = 53
    ElseIf strListGenre = "Eurodance" Then
        Genre = 54
    ElseIf strListGenre = "Dream" Then
        Genre = 55
    ElseIf strListGenre = "Southern Rock" Then
        Genre = 56
    ElseIf strListGenre = "Comedy" Then
        Genre = 57
    ElseIf strListGenre = "Cult" Then
        Genre = 58
    ElseIf strListGenre = "Gangsta" Then
        Genre = 59
    ElseIf strListGenre = "Top 40" Then
        Genre = 60
    ElseIf strListGenre = "Christian Rap" Then
        Genre = 61
    ElseIf strListGenre = "Pop/Funk" Then
        Genre = 62
    ElseIf strListGenre = "Jungle" Then
        Genre = 63
    ElseIf strListGenre = "Native American" Then
        Genre = 64
    ElseIf strListGenre = "Cabaret" Then
        Genre = 65
    ElseIf strListGenre = "New Wave" Then
        Genre = 66
    ElseIf strListGenre = "Psychadelic" Then
        Genre = 67
    ElseIf strListGenre = "Rave" Then
        Genre = 68
    ElseIf strListGenre = "Showtunes" Then
        Genre = 69
    ElseIf strListGenre = "Trailer" Then
        Genre = 70
    ElseIf strListGenre = "Low-Fi" Then
        Genre = 71
    ElseIf strListGenre = "Tribal" Then
        Genre = 72
    ElseIf strListGenre = "Acid Punk" Then
        Genre = 73
    ElseIf strListGenre = "Acid Jazz" Then
        Genre = 74
    ElseIf strListGenre = "Polka" Then
        Genre = 75
    ElseIf strListGenre = "Retro" Then
        Genre = 76
    ElseIf strListGenre = "Musical" Then
        Genre = 77
    ElseIf strListGenre = "Rock & Roll" Then
        Genre = 78
    ElseIf strListGenre = "Hard Rock" Then
        Genre = 79
    ElseIf strListGenre = "Folk" Then
        Genre = 80
    ElseIf strListGenre = "Folk-Rock" Then
        Genre = 81
    ElseIf strListGenre = "National Folk" Then
        Genre = 82
    ElseIf strListGenre = "Swing" Then
        Genre = 83
    ElseIf strListGenre = "Fast Fusion" Then
        Genre = 84
    ElseIf strListGenre = "Bebop" Then
        Genre = 85
    ElseIf strListGenre = "Latin" Then
        Genre = 86
    ElseIf strListGenre = "Revival" Then
        Genre = 87
    ElseIf strListGenre = "Celtic" Then
        Genre = 88
    ElseIf strListGenre = "Bluegrass" Then
        Genre = 89
    ElseIf strListGenre = "Avantgarde" Then
        Genre = 90
    ElseIf strListGenre = "Gothic Rock" Then
        Genre = 91
    ElseIf strListGenre = "Progressive Rock" Then
        Genre = 92
    ElseIf strListGenre = "Psychadelic Rock" Then
        Genre = 93
    ElseIf strListGenre = "Symphonic Rock" Then
        Genre = 94
    ElseIf strListGenre = "Slow Rock" Then
        Genre = 95
    ElseIf strListGenre = "Big Band" Then
        Genre = 96
    ElseIf strListGenre = "Chorus" Then
        Genre = 97
    ElseIf strListGenre = "Easy Listening" Then
        Genre = 98
    ElseIf strListGenre = "Acoustic" Then
        Genre = 99
    ElseIf strListGenre = "Humour" Then
        Genre = 100
    ElseIf strListGenre = "Speech" Then
        Genre = 101
    ElseIf strListGenre = "Chanson" Then
        Genre = 102
    ElseIf strListGenre = "Opera" Then
        Genre = 103
    ElseIf strListGenre = "Chamber Music" Then
        Genre = 104
    ElseIf strListGenre = "Sonata" Then
        Genre = 105
    ElseIf strListGenre = "Symphony" Then
        Genre = 106
    ElseIf strListGenre = "Booty Bass" Then
        Genre = 107
    ElseIf strListGenre = "Primus" Then
        Genre = 108
    ElseIf strListGenre = "Porn Groove" Then
        Genre = 109
    ElseIf strListGenre = "Satire" Then
        Genre = 110
    ElseIf strListGenre = "Slow Jam" Then
        Genre = 111
    ElseIf strListGenre = "Club" Then
        Genre = 112
    ElseIf strListGenre = "Tango" Then
        Genre = 113
    ElseIf strListGenre = "Samba" Then
        Genre = 114
    ElseIf strListGenre = "Folklore" Then
        Genre = 115
    ElseIf strListGenre = "Ballad" Then
        Genre = 116
    ElseIf strListGenre = "Power Ballad" Then
        Genre = 117
    ElseIf strListGenre = "Rhythmic Soul" Then
        Genre = 118
    ElseIf strListGenre = "Freestyle" Then
        Genre = 119
    ElseIf strListGenre = "Duet" Then
        Genre = 120
    ElseIf strListGenre = "Punk Rock" Then
        Genre = 121
    ElseIf strListGenre = "Drum Solo" Then
        Genre = 122
    ElseIf strListGenre = "A Capella" Then
        Genre = 123
    ElseIf strListGenre = "Euro-House" Then
        Genre = 124
    ElseIf strListGenre = "Dance Hall" Then
        Genre = 125
    End If

Open strFilename For Binary Access Write As #1
Seek #1, FileLen(strFilename) - 127
Put #1, , Tag
Put #1, , Songname
Put #1, , Artist
Put #1, , Album
Put #1, , Year
Put #1, , Comment
Put #1, , Genre
Close #1

End Sub


Sub getgenrefromID(strListGenre As Integer)

'Populate the ID3v1 genre combo box


If strListGenre = 0 Then
    cboGenre1.Text = "Blues"
ElseIf strListGenre = 1 Then
    cboGenre1.Text = "Classic Rock"
ElseIf strListGenre = 2 Then
    cboGenre1.Text = "Country"
ElseIf strListGenre = 3 Then
    cboGenre1.Text = "Dance"
ElseIf strListGenre = 4 Then
    cboGenre1.Text = "Disco"
ElseIf strListGenre = 5 Then
    cboGenre1.Text = "Funk"
ElseIf strListGenre = 6 Then
    cboGenre1.Text = "Grunge"
ElseIf strListGenre = 7 Then
    cboGenre1.Text = "Hip-Hop"
ElseIf strListGenre = 8 Then
    cboGenre1.Text = "Jazz"
ElseIf strListGenre = 9 Then
    cboGenre1.Text = "Metal"
ElseIf strListGenre = 10 Then
    cboGenre1.Text = "New Age"
ElseIf strListGenre = 11 Then
    cboGenre1.Text = "Oldies"
ElseIf strListGenre = 12 Then
    cboGenre1.Text = "Other"
ElseIf strListGenre = 13 Then
    cboGenre1.Text = "Pop"
ElseIf strListGenre = 14 Then
    cboGenre1.Text = "R&B"
ElseIf strListGenre = 15 Then
    cboGenre1.Text = "Rap"
ElseIf strListGenre = 16 Then
    cboGenre1.Text = "Reggae"
ElseIf strListGenre = 17 Then
    cboGenre1.Text = "Rock"
ElseIf strListGenre = 18 Then
    cboGenre1.Text = "Techno"
ElseIf strListGenre = 19 Then
    cboGenre1.Text = "Industrial"
ElseIf strListGenre = 20 Then
    cboGenre1.Text = "Alternative"
ElseIf strListGenre = 21 Then
    cboGenre1.Text = "Ska"
ElseIf strListGenre = 22 Then
    cboGenre1.Text = "Death Metal"
ElseIf strListGenre = 23 Then
    cboGenre1.Text = "Pranks"
ElseIf strListGenre = 24 Then
    cboGenre1.Text = "Soundtrack"
ElseIf strListGenre = 25 Then
    cboGenre1.Text = "Euro-Techno"
ElseIf strListGenre = 26 Then
    cboGenre1.Text = "Ambient"
ElseIf strListGenre = 27 Then
    cboGenre1.Text = "Trip-Hop"
ElseIf strListGenre = 28 Then
    cboGenre1.Text = "Vocal"
ElseIf strListGenre = 29 Then
    cboGenre1.Text = "Jazz+Funk"
ElseIf strListGenre = 30 Then
    cboGenre1.Text = "Fusion"
ElseIf strListGenre = 31 Then
    cboGenre1.Text = "Trance"
ElseIf strListGenre = 32 Then
    cboGenre1.Text = "Classical"
ElseIf strListGenre = 33 Then
    cboGenre1.Text = "Instrumental"
ElseIf strListGenre = 34 Then
    cboGenre1.Text = "Acid"
ElseIf strListGenre = 35 Then
    cboGenre1.Text = "House"
ElseIf strListGenre = 36 Then
    cboGenre1.Text = "Game"
ElseIf strListGenre = 37 Then
    cboGenre1.Text = "Sound Clip"
ElseIf strListGenre = 38 Then
    cboGenre1.Text = "Gospel"
ElseIf strListGenre = 39 Then
    cboGenre1.Text = "Noise"
ElseIf strListGenre = 40 Then
    cboGenre1.Text = "AlternRock"
ElseIf strListGenre = 41 Then
    cboGenre1.Text = "Bass"
ElseIf strListGenre = 42 Then
    cboGenre1.Text = "Soul"
ElseIf strListGenre = 43 Then
    cboGenre1.Text = "Punk"
ElseIf strListGenre = 44 Then
    cboGenre1.Text = "Space"
ElseIf strListGenre = 45 Then
    cboGenre1.Text = "Meditative"
ElseIf strListGenre = 46 Then
    cboGenre1.Text = "Instrumental Pop"
ElseIf strListGenre = 47 Then
    cboGenre1.Text = "Instrumental Rock"
ElseIf strListGenre = 48 Then
    cboGenre1.Text = "Ethnic"
ElseIf strListGenre = 49 Then
    cboGenre1.Text = "Gothic"
ElseIf strListGenre = 50 Then
    cboGenre1.Text = "Darkwave"
ElseIf strListGenre = 51 Then
    cboGenre1.Text = "Techno-Indistrial"
ElseIf strListGenre = 52 Then
    cboGenre1.Text = "Electronic"
ElseIf strListGenre = 53 Then
    cboGenre1.Text = "Pop-Folk"
ElseIf strListGenre = 54 Then
    cboGenre1.Text = "Eurodance"
ElseIf strListGenre = 55 Then
    cboGenre1.Text = "Dream"
ElseIf strListGenre = 56 Then
    cboGenre1.Text = "Southern Rock"
ElseIf strListGenre = 57 Then
    cboGenre1.Text = "Comedy"
ElseIf strListGenre = 58 Then
    cboGenre1.Text = "Cult"
ElseIf strListGenre = 59 Then
    cboGenre1.Text = "Gansta"
ElseIf strListGenre = 60 Then
    cboGenre1.Text = "Top 40"
ElseIf strListGenre = 61 Then
    cboGenre1.Text = "Christian Rap"
ElseIf strListGenre = 62 Then
    cboGenre1.Text = "Pop/Funk"
ElseIf strListGenre = 63 Then
    cboGenre1.Text = "Jungle"
ElseIf strListGenre = 64 Then
    cboGenre1.Text = "Native American"
ElseIf strListGenre = 65 Then
    cboGenre1.Text = "Cabaret"
ElseIf strListGenre = 66 Then
    cboGenre1.Text = "New Wave"
ElseIf strListGenre = 67 Then
    cboGenre1.Text = "Psychaledic"
ElseIf strListGenre = 68 Then
    cboGenre1.Text = "Rave"
ElseIf strListGenre = 69 Then
    cboGenre1.Text = "Showtunes"
ElseIf strListGenre = 70 Then
    cboGenre1.Text = "Trailer"
ElseIf strListGenre = 71 Then
    cboGenre1.Text = "Lo-Fi"
ElseIf strListGenre = 72 Then
    cboGenre1.Text = "Tribal"
ElseIf strListGenre = 73 Then
    cboGenre1.Text = "Acid Punk"
ElseIf strListGenre = 74 Then
    cboGenre1.Text = "Acid Jazz"
ElseIf strListGenre = 75 Then
    cboGenre1.Text = "Polka"
ElseIf strListGenre = 76 Then
    cboGenre1.Text = "Retro"
ElseIf strListGenre = 77 Then
    cboGenre1.Text = "Musical"
ElseIf strListGenre = 78 Then
    cboGenre1.Text = "Rock & Roll"
ElseIf strListGenre = 79 Then
    cboGenre1.Text = "Hard Rock"
ElseIf strListGenre = 80 Then
    cboGenre1.Text = "Folk"
ElseIf strListGenre = 81 Then
    cboGenre1.Text = "Folk-Rock"
ElseIf strListGenre = 82 Then
    cboGenre1.Text = "National Folk"
ElseIf strListGenre = 83 Then
    cboGenre1.Text = "Swing"
ElseIf strListGenre = 84 Then
    cboGenre1.Text = "Fast Fusion"
ElseIf strListGenre = 85 Then
    cboGenre1.Text = "Bebop"
ElseIf strListGenre = 86 Then
    cboGenre1.Text = "Latin"
ElseIf strListGenre = 87 Then
    cboGenre1.Text = "Revival"
ElseIf strListGenre = 88 Then
    cboGenre1.Text = "Celtic"
ElseIf strListGenre = 89 Then
    cboGenre1.Text = "Bluegrass"
ElseIf strListGenre = 90 Then
    cboGenre1.Text = "Avantgarde"
ElseIf strListGenre = 91 Then
    cboGenre1.Text = "Gothic Rock"
ElseIf strListGenre = 92 Then
    cboGenre1.Text = "Progressive Rock"
ElseIf strListGenre = 93 Then
    cboGenre1.Text = "Psychadelic Rock"
ElseIf strListGenre = 94 Then
    cboGenre1.Text = "Symphonic Rock"
ElseIf strListGenre = 95 Then
    cboGenre1.Text = "Slow Rock"
ElseIf strListGenre = 96 Then
    cboGenre1.Text = "Big Band"
ElseIf strListGenre = 97 Then
    cboGenre1.Text = "Chorus"
ElseIf strListGenre = 98 Then
    cboGenre1.Text = "Easy Listening"
ElseIf strListGenre = 99 Then
    cboGenre1.Text = "Acoustic"
ElseIf strListGenre = 100 Then
    cboGenre1.Text = "Humour"
ElseIf strListGenre = 101 Then
    cboGenre1.Text = "Speech"
ElseIf strListGenre = 102 Then
    cboGenre1.Text = "Chanson"
ElseIf strListGenre = 103 Then
    cboGenre1.Text = "Opera"
ElseIf strListGenre = 104 Then
    cboGenre1.Text = "Chamber Music"
ElseIf strListGenre = 105 Then
    cboGenre1.Text = "Sonata"
ElseIf strListGenre = 106 Then
    cboGenre1.Text = "Symphony"
ElseIf strListGenre = 107 Then
    cboGenre1.Text = "Booty Bass"
ElseIf strListGenre = 108 Then
    cboGenre1.Text = "Primus"
ElseIf strListGenre = 109 Then
    cboGenre1.Text = "Porn Groove"
ElseIf strListGenre = 110 Then
    cboGenre1.Text = "Satire"
ElseIf strListGenre = 111 Then
    cboGenre1.Text = "Slow Jam"
ElseIf strListGenre = 112 Then
    cboGenre1.Text = "Club"
ElseIf strListGenre = 113 Then
    cboGenre1.Text = "Tango"
ElseIf strListGenre = 114 Then
    cboGenre1.Text = "Samba"
ElseIf strListGenre = 115 Then
    cboGenre1.Text = "Folklore"
ElseIf strListGenre = 116 Then
    cboGenre1.Text = "Ballad"
ElseIf strListGenre = 117 Then
    cboGenre1.Text = "Power Ballad"
ElseIf strListGenre = 118 Then
    cboGenre1.Text = "Rhythmic Soul"
ElseIf strListGenre = 119 Then
    cboGenre1.Text = "Freestyle"
ElseIf strListGenre = 120 Then
    cboGenre1.Text = "Duet"
ElseIf strListGenre = 121 Then
    cboGenre1.Text = "Punk Rock"
ElseIf strListGenre = 122 Then
    cboGenre1.Text = "Drum Solo"
ElseIf strListGenre = 123 Then
    cboGenre1.Text = "A Capella"
ElseIf strListGenre = 124 Then
    cboGenre1.Text = "Euro-House"
ElseIf strListGenre = 125 Then
    cboGenre1.Text = "Dance Hall"
End If

End Sub

