## Ribbon Customisation

#### 1. Adding custom Ribbon tabs and commands

The [Office RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor/releases/tag/v1.9.0) is recommended for Ribbon customisation.  

Download and install RibbonX following the instructions provided with it.  

Download the file RIBBON_2007.xml from this folder in preparation for use.  

Prepare a document by installing and testing module SERIAL_PORT_VBA first. 

Follow the [HowTo](How-To.md) instructions to install the RIBBON_2007.xml sample customisation file.

#### 2. Office Icons

A list of icons included with Office is available here [Microsoft Office Icon Gallery Download](https://www.microsoft.com/en-nz/download/confirmation.aspx?id=21103)

Further information can be found online by searching for *msoImage*

Ribbon Office icons can be changed by editing the required XML file section in RibbonX, e.g. `imageMso="NewOfficeIconName"` 

#### 3. Custom Icons

Custom icons can also be added from RibbonX. Use the **Insert > Icons** menu option to add a new icon to the document. 

Ribbon Custom icons can be changed by editing the required XML file section in RibbonX, e.g. `image="MyPrivateIconName"` 

The following image filetypes are can be used. Image size should be between 16*16 to 128*128 

 .bmp  
 .gif   
 .jpg  
 .png  

Check online for further information on supported icon types and sizes for your Office version.

#### VBA Development

Further VBA development is required to make the document suitable for your usage. 


