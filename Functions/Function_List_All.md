## VBA Function List

Main user-defined functions are as follows. 


First parameter (1) is a valid COM Port number on host.

| VBA Function                         | Description                                                                                                   |
| ------------------------------------ | --------------------------------------------------------------------------------------------------------------|
| `debug_com_port(1)`                  | Toggles debug messaging on/off (debug results in VBA Immediate window)                                        |
| `debug_com_port(1,True)`             | Set port debug messaging on                                                                                   |          
| `debug_com_port(1,False)`            | Set port debug messaging off                                                                                  |  
| `start_com_port(1)`                  | Starts port with existing settings. Returns `True` if successful, `False` if start fails for any reason.      | 
| `start_com_port(1,"Baud=1200")`      | Starts port with settings as supplied. Returns `True` or `False` as above.                                    |
| `start_com_port(1,SCANNER)`          | Starts port with settings defined in VBA string constant or variable e.g. SCANNER                             |
| `check_com_port(1)`                  | Returns number of input characters waiting to be read (no delay). Return value -1 indicates error.            |
| `wait_com_port(1)`                   | Wait for up to 333mS (default) before timing out. Returns `True` if receive data waiting.                     |
| `wait_com_port(1,500)`               | As above, can optionally specify wait time (500) in milliseconds. Timesliced to avoid VBA hanging.            |  
| `get_com_port(1)`                    | Receives a single character string from a started com port.                                                   |
| `put_com_port(1,"A")`                | Sends a single character to a started com port.                                                               |
| `read_com_port(1,20)`                | Reads up to specified number (20) of characters. No delay, max characters = approx 1 second timeslice.        |
| `send_com_port(1,V)`                 | Sends variable V. Function converts V to String and calls `transmit_com_port`.                                |
| `receive_com_port(1)`                | Receives all data from port, timesliced for low port speeds and/or large data transfers.                      |
| `transmit_com_port(1,"QWERTY")`      | Sends string to port, in timeslices of approx 1 second to avoid VBA 'not responding'                          |
| `device_ready(1)`                    | Returns `True` if port started and Data Set Ready (DSR) input signal active.                                  |
| `device_calling(1)`                  | Returns `True` if port started and Ring Indicate (RI) input signal active.                                    |
| `carrier_detect(1)`                  | Returns `True` if port started and Carrier Detect (CD) input signal active.                                   |
| `clear_to_send(1)`                   | Returns `True` if port started and Clear To Send (CTS) input signal active.                                   |
| `request_to_send(1,[1/0])`           | Sets Request To Send (RTS) output signal on/off `1/0` Returns `True` if port started and RTS set/cleared      |
| `signal_com_port(1,signal)`          | Set/clear Break, DTR, RTS output signals, see - [Escape Comm Function signal values](https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-escapecommfunction)
| `stop_com_port(1)`                   | Stops port and hands its control back to Windows.                                                             |

Other Public functions such as `show_port_errors(1)` etc. should only be used in the Immediate window for further information if required.
Private functions are not intended to be called directly by users.
