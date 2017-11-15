Two custom UserControls for use on Userforms from Microsoft Visual Forms 2.0 in VBA

Both controls use an MSForms.Image class as the drawing surface and is used to visualize
the difference between two different colors in CIELAB colorspace.

cLabScope shows how the a* and b* component from one color relates to the other color
cLScope show how the L* component from one color relates to the other color


How to use this
--------------------------------------------------------
Import the following files into a new VBA project
* cLScope.cls
* cLabScope
* factory.bas
* ColorConversions (from https://github.com/arildj78/ColorConversions) 

Create a new userform and populate it with two image usercontrols

Add the following code to the Userform source:

    Public Scope1 As cLabScope
    Public Scope2 As cLScope
    
    Private Sub UserForm_Initialize()
        Set Scope1 = factory.Create_LabScope(Me, Image1)
        Set Scope2 = factory.Create_LScope(Me, Image2)
    End Sub

