# Ribbon Customisation in Access


Steps to customise the Ribbon in Access only are summarised below

1. Open new or existing Access document
2. Check if a database table **`USysRibbons`** exists first
3. If not, then create a new table **`USysRibbons`** 
4. Table Columns = `ID (AutoNumber) , RibbonName (Short Text) , RibbonXml (Long Text)`
5. Insert a new row into table `USysRibbons` , RibbonName = "COM Port 1"
6. Update the new row by copying the file `RIBBON_ACCESS_2007.xml` into column RibbonXML
7. 
8. 
  
