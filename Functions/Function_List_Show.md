# VBA Show Functions

## Show COM port information

First parameter (1) is a valid[^1] COM Port number on host PC

| VBA Function                         | Description                                                                                                   |
| ------------------------------------ | --------------------------------------------------------------------------------------------------------------|
| `show_port_dcb(1)`                   | Display contents of the port's Device Control Block                                                           |
| `show_port_errors(1)`                | Show port overflow/overrun/parity/framing/break conditions                                                    |          
| `show_port_modem(1)`                 | Show port modem signals DSR/CTS/RING/CD                                                                       |  
| `show_port_queues(1)`                | Show port input and output data queues                                                                        |  
| `show_port_status(1)`                |                                                                          |
| `start_com_port(1,SCANNER)`          | Starts port with settings defined in string constant or variable e.g. SCANNER                                 |
| `stop_com_port(1)`                   | Stops port and hands its control back to Windows                                                              |

* Show results are in the VBA Immediate Window (Control-G)
* Debug functions return `True` or `False` to indicate debug state
* Other functions return `True` or `False` to indicate success or failure

[^1]: Valid Minimum and Maximum port numbers should be defined in declarations section at the start of the module. 
  

