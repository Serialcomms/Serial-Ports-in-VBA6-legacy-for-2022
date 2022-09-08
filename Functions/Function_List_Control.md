1
First parameter (1) is a valid COM Port number on host.

| VBA Function                         | Description                                                                                                   |
| ------------------------------------ | --------------------------------------------------------------------------------------------------------------|
| `debug_com_port(1)`                  | Toggles debug messaging on/off (debug results in VBA Immediate window)                                        |
| `debug_com_port(1,True)`             | Set port debug messaging on                                                                                   |          
| `debug_com_port(1,False)`            | Set port debug messaging off                                                                                  |  
| `start_com_port(1)`                  | Starts port with existing settings. Returns `True` if successful, `False` if start fails for any reason.      | 
| `start_com_port(1,"Baud=1200")`      | Starts port with settings as supplied. Returns `True` or `False` as above.                                    |
| `start_com_port(1,SCANNER)`          | Starts port with settings defined in VBA string constant or variable e.g. SCANNER                             |
| `stop_com_port(1)`                   | Stops port and hands its control back to Windows.                                                             |

Other Public functions such as `show_port_errors(1)` etc. should only be used in the Immediate window for further information if required.
Private functions are not intended to be called directly by users.