Attribute VB_Name = "factory"

Public Function Create_LabScope(ByRef Parent As Object, _
                                ByRef DisplaySurface As MSForms.Image, _
                                Optional HitLabA As Double = 0#, _
                                Optional HitLabB As Double = 0#, _
                                Optional BackColor As Long = &H8000000F, _
                                Optional ForeColor As Long = &H80000012, _
                                Optional ReticleColor As Long = &HFF7F00, _
                                Optional Spokes As Long = 4&, _
                                Optional Rings As Long = 3&, _
                                Optional Padding As Long = 0&, _
                                Optional UnitsPrDivision As Double = 5#, _
                                Optional TgtLabL As Double = 60#, _
                                Optional TgtLabA As Double = 0#, _
                                Optional TgtLabB As Double = 0#, _
                                Optional LabExaggeration As Double = 5#, _
                                Optional ScrollBeyondLimit As Boolean = True, _
                                Optional ColorSpace As tColorSpace = tColorSpace.AdobeRGB, _
                                Optional MaskDeltaE As Single = 5#, _
                                Optional OutOfGamutWarning As Boolean = False) As cLabScope
                                
    
    Set Create_LabScope = New cLabScope
    Create_LabScope.InitiateProperties Parent:=Parent, _
                                       DisplaySurface:=DisplaySurface, _
                                       HitLabA:=HitLabA, _
                                       HitLabB:=HitLabB, _
                                       BackColor:=BackColor, _
                                       ForeColor:=ForeColor, _
                                       ReticleColor:=ReticleColor, _
                                       Spokes:=Spokes, _
                                       Rings:=Rings, _
                                       Padding:=Padding, _
                                       UnitsPrDivision:=UnitsPrDivision, _
                                       TgtLabL:=TgtLabL, _
                                       TgtLabA:=TgtLabA, _
                                       TgtLabB:=TgtLabB, _
                                       LabExaggeration:=LabExaggeration, _
                                       ScrollBeyondLimit:=ScrollBeyondLimit, _
                                       ColorSpace:=ColorSpace, _
                                       MaskDeltaE:=MaskDeltaE, _
                                       OutOfGamutWarning:=OutOfGamutWarning
                                    
End Function


Public Function Create_LScope(ByRef Parent As Object, _
                             ByRef DisplaySurface As MSForms.Image, _
                             Optional HitLabL As Double = 60#, _
                             Optional ForeColor As Long = &H80000012, _
                             Optional ReticleColor As Long = &HFF7F00, _
                             Optional Padding As Long = 5&, _
                             Optional Divisions As Long = 6&, _
                             Optional UnitsPrDivision As Double = 5#, _
                             Optional TgtLabL As Double = 60#, _
                             Optional TgtLabA As Double = 0#, _
                             Optional TgtLabB As Double = 0#, _
                             Optional LabExaggeration As Double = 2#, _
                             Optional ScrollBeyondLimit As Boolean = True) As cLScope
    
    Set Create_LScope = New cLScope
    Create_LScope.InitiateProperties Parent:=Parent, _
                                    DisplaySurface:=DisplaySurface, _
                                    HitLabL:=HitLabL, _
                                    ForeColor:=ForeColor, _
                                    ReticleColor:=ReticleColor, _
                                    Padding:=Padding, _
                                    Divisions:=Divisions, _
                                    UnitsPrDivision:=UnitsPrDivision, _
                                    TgtLabL:=TgtLabL, _
                                    TgtLabA:=TgtLabA, _
                                    TgtLabB:=TgtLabB, _
                                    LabExaggeration:=LabExaggeration, _
                                    ScrollBeyondLimit:=ScrollBeyondLimit
                                    
End Function

Public Function Create_ColorTile(ByRef Parent As Object, _
                                 ByRef DisplaySurface As MSForms.Image, _
                                 Optional LabScope As cLabScope = Nothing, _
                                 Optional LScope As cLScope = Nothing, _
                                 Optional TgtLabL As Double = 60#, _
                                 Optional TgtLabA As Double = 0#, _
                                 Optional TgtLabB As Double = 0#, _
                                 Optional HitLabL As Double = 50#, _
                                 Optional HitLabA As Double = 0#, _
                                 Optional HitLabB As Double = 0#, _
                                 Optional Frame As Boolean = False, _
                                 Optional FrameThickness As Long = 1&, _
                                 Optional xOut As Boolean = False, _
                                 Optional ForeColor As Long = &H80000012) As cColorTile
    
    Set Create_ColorTile = New cColorTile
    Create_ColorTile.InitiateProperties Parent:=Parent, _
                                        DisplaySurface:=DisplaySurface, _
                                        LabScope:=LabScope, _
                                        LScope:=LScope, _
                                        TgtLabL:=TgtLabL, _
                                        TgtLabA:=TgtLabA, _
                                        TgtLabB:=TgtLabB, _
                                        HitLabL:=HitLabL, _
                                        HitLabA:=HitLabA, _
                                        HitLabB:=HitLabB, _
                                        Frame:=Frame, _
                                        FrameThickness:=FrameThickness, _
                                        xOut:=xOut, _
                                        ForeColor:=ForeColor

End Function
