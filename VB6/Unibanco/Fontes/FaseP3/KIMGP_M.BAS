Attribute VB_Name = "KOFAX_IMGP"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        Copyright (c) 1993-1996, Kofax Image Products. All rights reserved.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'       File: kimgp_m.bas
'    Purpose: Visual Basic Constants for use with KIMGP.VBX custom control
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' tab=4 '

Option Explicit


' Action Settings
Global Const KIACTIONNONE = 0
Global Const KIACTIONOPENDIALOG = 1
Global Const KIACTIONREADINI = 2
Global Const KIACTIONWRITEINI = 3

' ActiveDialog Settings
Global Const KIDIALOGABOUT = 0
Global Const KIDIALOGBARCODE = 1
Global Const KIDIALOGANNOTATE = 2
Global Const KIDIALOGPATCHCODE = 3
Global Const KIDIALOGPATCHCODEADVANCED = 4
Global Const KIDIALOGDESKEW = 5
Global Const KIDIALOGBLACKBORDER = 6
Global Const KIDIALOGDESHADE = 7
Global Const KIDIALOGDESPECKLE = 8
Global Const KIDIALOGEDGEENHANCEMENT = 9
Global Const KIDIALOGLINEREMOVAL = 10
Global Const KIDIALOGSTREAKREMOVAL = 11

' PatchTriggers Settings
Global Const KIPATCHNONE = &H0
Global Const KIPATCH3 = &H1
Global Const KIPATCH2 = &H2
Global Const KIPATCHT = &H4
Global Const KIPATCH1 = &H8
Global Const KIPATCH4 = &H10
Global Const KIPATCH6 = &H20
Global Const KIPATCHTRIGGERSALLCAP = &H3F

' BarDensity Settings
Global Const KIBARDENSITYNORMAL = 0
Global Const KIBARDENSITYDARK = 1
Global Const KIBARDENSITYLIGHT = 2

' BarOrientation Settings
Global Const KIBARORIENTATION0 = &H1
Global Const KIBARORIENTATION90 = &H2
Global Const KIBARORIENTATION180 = &H4
Global Const KIBARORIENTATION270 = &H8
Global Const KIBARORIENTATIONALLCAP = &HF

' BarQuality Settings
Global Const KIBARQUALITYGOOD = 0
Global Const KIBARQUALITYNORMAL = 1
Global Const KIBARQUALITYPOOR = 2

' BarSkew Settings
Global Const KIBARSKEWNONE = 0
Global Const KIBARSKEWLIGHT = 1
Global Const KIBARSKEWHEAVY = 2

' BarType Settings
Global Const KIBARTYPE3OF9 = &H1
Global Const KIBARTYPE2OF5 = &H2
Global Const KIBARTYPE128 = &H4
Global Const KIBARTYPEEAN = &H8
Global Const KIBARTYPEUPCA = &H10
Global Const KIBARTYPEUPCE = &H20
Global Const KIBARTYPECODABAR = &H40
Global Const KIBARTYPECODE93 = &H80
Global Const KIBARTYPEPOSTNET = &H100
Global Const KIBARTYPELINEAR2OF5 = &H200
Global Const KIBARTYPEALLCAP = &H3FF

' EdgeEnhancementAlgorithm Settings
Global Const KIEDGECHARSMOOTHING = &H1
Global Const KIEDGETHINNING = &H2
Global Const KIEDGETHICKEN = &H4
Global Const KIEDGESMOOTHCLEAN = &H8
Global Const KIEDGEFILLLINEFILT = &H10
Global Const KIEDGEPRESERVESMOOTHCLEAN = &H20
Global Const KIEDGEPRESERVEFILLLINE = &H40
Global Const KIEDGELIGHTTHICKEN = &H80
Global Const KIEDGEOUTLINE = &H100
Global Const KIEDGEALGORITHMALLCAP = &H1FF

' SkewMinAngle Settings
Global Const KISKEWMINANGLE = 0
Global Const KISKEWMAXANGLE = 12#

' FontName Settings
Global Const KIFONTNAMESANSSERIF = 0
Global Const KIFONTNAMESERIF = 1

' FontSize Settings
Global Const KIFONTSIZE8 = 0
Global Const KIFONTSIZE10 = 1
Global Const KIFONTSIZE12 = 2

' FontDpi Settings
Global Const KIFONTDPI200 = 0
Global Const KIFONTDPI300 = 1

' FontOrientation Settings
Global Const KIFONTORIENTATIONHORZ0 = 0
Global Const KIFONTORIENTATIONVERT0 = 1
Global Const KIFONTORIENTATIONVERT90 = 2
Global Const KIFONTORIENTATIONVERT270 = 3

' FontBackground Settings
Global Const KIFONTBACKGROUNDTRANSPARENT = 0
Global Const KIFONTBACKGROUNDOPAQUE = 1

' DespeckleWidth Settings
Global Const KIDESPECKLEWIDTHMINPIXELS = 1
Global Const KIDESPECKLEWIDTHMAXPIXELS = 32
Global Const KIDESPECKLEHEIGHTMINPIXELS = 1
Global Const KIDESPECKLEHEIGHTMAXPIXELS = 32

' BlackBorderWhiteNoiseGap Settings
Global Const KIBLACKBORDERGAPMINPIXELS = 0
Global Const KIBLACKBORDERGAPMAXPIXELS = 10

' StreakWidth Settings
Global Const KISTREAKWIDTHMINPIXELS = 0
Global Const KISTREAKWIDTHMAXPIXELS = 10

' DeshadeVertSpeckleAdj Settings
Global Const KIDESHADEMINWIDTH = 50
Global Const KIDESHADEMAXWIDTH = &HFFFF&
Global Const KIDESHADEMINHEIGHT = 20
Global Const KIDESHADEMAXHEIGHT = &HFFFF&
Global Const KIDESHADESPECKLEMIN = 1
Global Const KIDESHADESPECKLEMAX = 12
Global Const KIDESHADESPECKLEMINADJ = -3
Global Const KIDESHADESPECKLEMAXADJ = 3
Global Const KILINELENGTHMIN = 0
Global Const KILINELENGTHMAX = 32767
Global Const KILINEHEIGHTMIN = 0
Global Const KILINEHEIGHTMAX = 32767
Global Const KILINEWIDTHMIN = 0
Global Const KILINEWIDTHMAX = 32767
Global Const KILINEGAPMIN = 0
Global Const KILINEGAPMAX = 10
Global Const KILINERECONSTRUCTMIN = 0
Global Const KILINERECONSTRUCTMAX = 10
Global Const KILINEEDGECLEANMIN = 0
Global Const KILINEEDGECLEANMAX = 10
