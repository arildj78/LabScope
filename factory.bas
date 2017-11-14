Attribute VB_Name = "factory"

Public Function Create_LabScope(ByRef Parent As Object, _
                                ByRef DisplaySurface As MSForms.Image, _
                                Optional HitLabA = 0#, _
                                Optional HitLabB = 0#, _
                                Optional BackColor = &H8000000F, _
                                Optional ForeColor = &H80000012, _
                                Optional ReticleColor = &HFF7F00, _
                                Optional Spokes = 4, _
                                Optional Rings = 3, _
                                Optional Padding = 0, _
                                Optional UnitsPrDivision = 5#, _
                                Optional TgtLabL = 60#, _
                                Optional TgtLabA = 0#, _
                                Optional TgtLabB = 0#, _
                                Optional LabExaggeration = 5#) As cLabScope
    
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
                                       LabExaggeration:=LabExaggeration
                                    
End Function


Public Function Create_LScope(ByRef Parent As Object, _
                             ByRef DisplaySurface As MSForms.Image, _
                             Optional HitLabL = 60#, _
                             Optional BackColor = &H8000000F, _
                             Optional ForeColor = &H80000012, _
                             Optional ReticleColor = &HFF7F00, _
                             Optional Padding = 5, _
                             Optional Divisions = 6, _
                             Optional UnitsPrDivision = 5#, _
                             Optional TgtLabL = 60#, _
                             Optional TgtLabA = 0#, _
                             Optional TgtLabB = 0#, _
                             Optional LabExaggeration = 2#) As cLScope
    
    Set Create_LScope = New cLScope
    Create_LScope.InitiateProperties Parent:=Parent, _
                                    DisplaySurface:=DisplaySurface, _
                                    HitLabL:=HitLabL, _
                                    BackColor:=BackColor, _
                                    ForeColor:=ForeColor, _
                                    ReticleColor:=ReticleColor, _
                                    Padding:=Padding, _
                                    Divisions:=Divisions, _
                                    UnitsPrDivision:=UnitsPrDivision, _
                                    TgtLabL:=TgtLabL, _
                                    TgtLabA:=TgtLabA, _
                                    TgtLabB:=TgtLabB, _
                                    LabExaggeration:=LabExaggeration
                                    
End Function


