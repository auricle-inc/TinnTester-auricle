Attribute VB_Name = "modTinTest"
Public boolDblClick As Boolean 'variable to prevent users from double clicking
Public vRes As Integer 'vertical resolution.  Change before compiling to setup program for 1024 resolution
Public usePA52 As Boolean
Public CalibData(1 To 4, 1 To 11) As Single
'CalibData(1, 1 to 11) holds dBSPL levsl for ringing sounds when PA5 = 0
'CalibData(2, 1 to 11) holds dBSPL levsl for pure tone sounds when PA5 = 0
'CalibData(3, 1 to 11) holds dBSPL levsl for hissing sounds when PA5 = 0
'calibdata(4,1) = dBSPL levsl for WN when PA5 = 0
Public UserName As String
Public UserCity As String
Public UserProv As String
Public UserCountry As String
Public UserAge As String
Public UserSex As String
Public UserOnset As String
Public TLoud As Integer 'tinnitus loudness rating
Public English As Boolean 'true = in english, false = in french
Public RI5k As Single  'will hold calculated RI value at 5khz to pass into report function. Calculated in step 9
Public UserTL As String 'will hold user location of tinnitus for report puproses only
Public UserSorP As String 'will hold user Steady or Pulsing property of tinnitus for report puproses only
Public UserBW As String ' will hold users bandwitdh (ringing/hissing/pure) for report puproses only
Public ActiveLock As ActiveLock3.IActiveLock
Public WorkingDir As String
Public WorkingFile As String
Public PReport As Boolean
Public TinTrainComplete As Boolean
Public OneStep As Boolean ' this will make sure user selects AT LEAST one sound to listen too

