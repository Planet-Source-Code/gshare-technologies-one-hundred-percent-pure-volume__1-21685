Attribute VB_Name = "BasSetVolume"
      
      Public Const SETVOLMMYSERR_NOERROR = 0
      Public Const SETVOLMAXPNAMELEN = 32
      Public Const SETVOLMIXER_LONG_NAME_CHARS = 64
      Public Const SETVOLMIXER_SHORT_NAME_CHARS = 16
      Public Const SETVOLMIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
      Public Const SETVOLMIXER_GETCONTROLDETAILSF_VALUE = &H0&
      Public Const SETVOLMIXER_GETLINECONTROLSF_ONEBYTYPE = &H2&
      Public Const SETVOLSETVOLMIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
      Public Const SETVOLSETVOLMIXERLINE_COMPONENTTYPE_SRC_FIRST = &H1000&
      
      Public Const SETVOLSETVOLMIXERLINE_COMPONENTTYPE_DST_SPEAKERS = _
                     (SETVOLSETVOLMIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
                     
      Public Const SETVOLSETVOLMIXERLINE_COMPONENTTYPE_SRC_MICROPHONE = _
                     (SETVOLSETVOLMIXERLINE_COMPONENTTYPE_SRC_FIRST + 3)
      
      Public Const SETVOLSETVOLMIXERLINE_COMPONENTTYPE_SRC_LINE = _
                     (SETVOLSETVOLMIXERLINE_COMPONENTTYPE_SRC_FIRST + 2)
      
      Public Const SETVOLMIXERCONTROL_CT_CLASS_FADER = &H50000000
      Public Const SETVOLMIXERCONTROL_CT_UNITS_UNSIGNED = &H30000
      
      Public Const SETVOLMIXERCONTROL_CONTROLTYPE_FADER = _
                     (SETVOLMIXERCONTROL_CT_CLASS_FADER Or _
                     SETVOLMIXERCONTROL_CT_UNITS_UNSIGNED)
      
      Public Const SETVOLMIXERCONTROL_CONTROLTYPE_VOLUME = _
                     (SETVOLMIXERCONTROL_CONTROLTYPE_FADER + 1)
      
      Private Declare Function SETVOLMIXERClose Lib "winmm.dll" _
                     (ByVal hmx As Long) As Long
         
      Private Declare Function mixerGetControlDetails Lib "winmm.dll" _
                     Alias "mixerGetControlDetailsA" _
                     (ByVal hmxobj As Long, _
                     psetvolmxcd As SETVOLMIXERCONTROLDETAILS, _
                     ByVal fdwDetails As Long) As Long
         
      Private Declare Function mixerGetDevCaps Lib "winmm.dll" _
                     Alias "mixerGetDevCapsA" _
                     (ByVal uMxId As Long, _
                     ByVal pmxcaps As SETVOLMIXERCAPS, _
                     ByVal cbmxcaps As Long) As Long
         
      Private Declare Function mixerGetID Lib "winmm.dll" _
                     (ByVal hmxobj As Long, _
                     pumxID As Long, _
                     ByVal fdwId As Long) As Long
                     
      Private Declare Function mixerGetLineControls Lib "winmm.dll" _
                     Alias "mixerGetLineControlsA" _
                     (ByVal hmxobj As Long, _
                     pmxlc As SETVOLMIXERLINECONTROLS, _
                     ByVal fdwControls As Long) As Long
                     
      Private Declare Function mixerGetLineInfo Lib "winmm.dll" _
                     Alias "mixerGetLineInfoA" _
                     (ByVal hmxobj As Long, _
                     pmxl As SETVOLMIXERLINE, _
                     ByVal fdwInfo As Long) As Long
                     
      Private Declare Function mixerGetNumDevs Lib "winmm.dll" () As Long
      
     
      Private Declare Function mixerMessage Lib "winmm.dll" _
                     (ByVal hmx As Long, _
                     ByVal uMsg As Long, _
                     ByVal dwParam1 As Long, _
                     ByVal dwParam2 As Long) As Long
                     
      Private Declare Function mixerOpen Lib "winmm.dll" _
                     (phmx As Long, _
                     ByVal uMxId As Long, _
                     ByVal dwCallback As Long, _
                     ByVal dwInstance As Long, _
                     ByVal fdwOpen As Long) As Long
                     
      Private Declare Function mixerSetControlDetails Lib "winmm.dll" _
                     (ByVal hmxobj As Long, _
                     psetvolmxcd As SETVOLMIXERCONTROLDETAILS, _
                     ByVal fdwDetails As Long) As Long
                     
      Declare Sub CopyStructFromPtr Lib "kernel32" _
                     Alias "RtlMoveMemory" _
                     (struct As Any, _
                     ByVal ptr As Long, ByVal cb As Long)
                     
      Declare Sub CopyPtrFromStruct Lib "kernel32" _
                     Alias "RtlMoveMemory" _
                     (ByVal ptr As Long, _
                     struct As Any, _
                     ByVal cb As Long)
                     
      Private Declare Function GlobalAlloc Lib "kernel32" _
                     (ByVal wFlags As Long, _
                     ByVal dwBytes As Long) As Long
                     
      Private Declare Function GlobalLock Lib "kernel32" _
                     (ByVal hmem As Long) As Long
                     
      Private Declare Function GlobalFree Lib "kernel32" _
                     (ByVal hmem As Long) As Long
      
      Type SETVOLMIXERCAPS
         wMid As Integer                   '  manufacturer id
         wPid As Integer                   '  product id
         vDriverVersion As Long            '  version of the driver
         szPname As String * SETVOLMAXPNAMELEN   '  product name
         fdwSupport As Long                '  misc. support bits
         cDestinations As Long             '  count of destinations
      End Type
      
      Type SETVOLMIXERCONTROL
         cbStruct As Long           '  size in Byte of SETVOLMIXERCONTROL
         dwControlID As Long        '  unique control id for mixer device
         dwControlType As Long      '  SETVOLMIXERCONTROL_CONTROLTYPE_xxx
         fdwControl As Long         '  SETVOLMIXERCONTROL_CONTROLF_xxx
         cMultipleItems As Long     '  if SETVOLMIXERCONTROL_CONTROLF_MULTIPLE set
         szShortName As String * SETVOLMIXER_SHORT_NAME_CHARS  ' short name of control
         szName As String * SETVOLMIXER_LONG_NAME_CHARS        ' long name of control
         lMinimum As Long           '  Minimum value
         lMaximum As Long           '  Maximum value
         reserved(10) As Long       '  reserved structure space
         End Type
      
      Type SETVOLMIXERCONTROLDETAILS
         cbStruct As Long       '  size in Byte of SETVOLMIXERCONTROLDETAILS
         dwControlID As Long    '  control id to get/set details on
         cChannels As Long      '  number of channels in paDetails array
         item As Long           '  hwndOwner or cMultipleItems
         cbDetails As Long      '  size of _one_ details_XX struct
         paDetails As Long      '  pointer to array of details_XX structs
      End Type
      
      Type SETVOLMIXERCONTROLDETAILS_UNSIGNED
         dwValue As Long        '  value of the control
      End Type
      
      Type SETVOLMIXERLINE
         cbStruct As Long               '  size of SETVOLMIXERLINE structure
         dwDestination As Long          '  zero based destination index
         dwSource As Long               '  zero based source index (if source)
         dwLineID As Long               '  unique line id for mixer device
         fdwLine As Long                '  state/information about line
         dwUser As Long                 '  driver specific information
         dwComponentType As Long        '  component type line connects to
         cChannels As Long              '  number of channels line supports
         cConnections As Long           '  number of connections (possible)
         cControls As Long              '  number of controls at this line
         szShortName As String * SETVOLMIXER_SHORT_NAME_CHARS
         szName As String * SETVOLMIXER_LONG_NAME_CHARS
         dwType As Long
         dwDeviceID As Long
         wMid  As Integer
         wPid As Integer
         vDriverVersion As Long
         szPname As String * SETVOLMAXPNAMELEN
      End Type
      
      Type SETVOLMIXERLINECONTROLS
         cbStruct As Long       '  size in Byte of SETVOLMIXERLINECONTROLS
         dwLineID As Long       '  line id (from SETVOLMIXERLINE.dwLineID)
                                '  SETVOLMIXER_GETLINECONTROLSF_ONEBYID or
         dwControl As Long      '  SETVOLMIXER_GETLINECONTROLSF_ONEBYTYPE
         cControls As Long      '  count of controls pmxctrl points to
         cbmxctrl As Long       '  size in Byte of _one_ SETVOLMIXERCONTROL
         pamxctrl As Long       '  pointer to first SETVOLMIXERCONTROL array
      End Type



    Public Const SETVOLMMYSERR_BASE = 0
    Public Const SETVOLMMYSERR_BADDEVICEID = (SETVOLMMYSERR_BASE + 2)

   
   Global SetVolHmixer As Long          ' mixer handle
   Global SetVolCtrl As SETVOLMIXERCONTROL ' waveout volume control
   Global SetMicCtrl As SETVOLMIXERCONTROL ' microphone volume control
   Global rc As Long              ' return code
   Global ok As Boolean           ' boolean return code


    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Function InitGetVolume()
    rc = mixerOpen(SetVolHmixer, 0, 0, 0, 0)
    If ((SETVOLMMYSERR_NOERROR <> rc)) Then
       InitGetVolume = False
    Exit Function
    End If
    
        ok = GetVolumeControl(SetVolHmixer, _
    SETVOLSETVOLMIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
    SETVOLMIXERCONTROL_CONTROLTYPE_VOLUME, _
    SetVolCtrl)
    
End Function


      Function GetVolumeControl(ByVal SetVolHmixer As Long, _
                              ByVal componentType As Long, _
                              ByVal ctrlType As Long, _
                              ByRef mxc As SETVOLMIXERCONTROL) As Boolean
                              
      ' This function attempts to obtain a mixer control. Returns True if successful.
         Dim mxlc As SETVOLMIXERLINECONTROLS
         Dim mxl As SETVOLMIXERLINE
         Dim hmem As Long

             
         mxl.cbStruct = Len(mxl)
         mxl.dwComponentType = componentType
      
         ' Obtain a line corresponding to the component type
         rc = mixerGetLineInfo(SetVolHmixer, mxl, SETVOLMIXER_GETLINEINFOF_COMPONENTTYPE)
         
         If (SETVOLMMYSERR_NOERROR = rc) Then
             mxlc.cbStruct = Len(mxlc)
             mxlc.dwLineID = mxl.dwLineID
             mxlc.dwControl = ctrlType
             mxlc.cControls = 1
             mxlc.cbmxctrl = Len(mxc)
             
             ' Allocate a buffer for the control
             hmem = GlobalAlloc(&H40, Len(mxc))
             mxlc.pamxctrl = GlobalLock(hmem)
             mxc.cbStruct = Len(mxc)
             
             ' Get the control
             rc = mixerGetLineControls(SetVolHmixer, _
                                       mxlc, _
                                       SETVOLMIXER_GETLINECONTROLSF_ONEBYTYPE)
                  
             If (SETVOLMMYSERR_NOERROR = rc) Then
                 GetVolumeControl = True
                 
                 ' Copy the control into the destination structure
                 CopyStructFromPtr mxc, mxlc.pamxctrl, Len(mxc)
             Else
                 GetVolumeControl = False
             End If
             GlobalFree (hmem)
             Exit Function
         End If
      
         GetVolumeControl = False
      End Function
      
      Function SetVolumeControl(ByVal SetVolHmixer As Long, _
                              mxc As SETVOLMIXERCONTROL, _
                              ByVal Volume As Long) As Boolean
      'This function sets the value for a volume control. Returns True if successful
                              
         Dim setvolmxcd As SETVOLMIXERCONTROLDETAILS
         Dim Vol As SETVOLMIXERCONTROLDETAILS_UNSIGNED
      
         setvolmxcd.item = 0
         setvolmxcd.dwControlID = mxc.dwControlID
         setvolmxcd.cbStruct = Len(setvolmxcd)
         setvolmxcd.cbDetails = Len(Vol)
         
         ' Allocate a buffer for the control value buffer
         hmem = GlobalAlloc(&H40, Len(Vol))
         setvolmxcd.paDetails = GlobalLock(hmem)
         setvolmxcd.cChannels = 1
         
         
         Vol.dwValue = Volume
         
         ' Copy the data into the control value buffer
         CopyPtrFromStruct setvolmxcd.paDetails, Vol, Len(Vol)
         
         ' Set the control value
         rc = mixerSetControlDetails(SetVolHmixer, _
                                    setvolmxcd, _
                                    SETVOLMIXER_SETCONTROLDETAILSF_VALUE)
         
         GlobalFree (hmem)
         If (SETVOLMMYSERR_NOERROR = rc) Then
             SetVolumeControl = True
         Else
             SetVolumeControl = False
         End If
      End Function

