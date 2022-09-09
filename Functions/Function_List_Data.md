### VBA functions to send, receive and check for data

First parameter (1) is a valid and Started COM Port number on host PC.

| VBA Function                    |  TS  | Description                                                                                                     |
| --------------------------------|:----:| --------------------------------------------------------------------------------------------------------------|
| `check_com_port(1)`             |  No  | Returns number of input characters waiting to be read (no delay). Return value -1 indicates error.            |
| `wait_com_port(1)`              |  Yes | Wait for up to 333mS (default) before timing out. Returns `True` if receive data waiting.                     |
| `wait_com_port(1,500)`          |  Yes | As above, can optionally specify wait time (500) in milliseconds.                                             |  
| `get_com_port(1)`               |   No | Receives a single character string from com port.                                                             |
| `put_com_port(1,"A")`           |   No | Sends a single character to com port.                                                                         |
| `read_com_port(1,20)`           |   No | Reads up to specified number (20) of characters. No delay, max characters read = approx 1 second timeslice.   |
| `send_com_port(1,V)`            |  Yes | Sends variable V. Function converts V to String and calls `transmit_com_port`.                                |
| `receive_com_port(1)`           |  Yes | Receives all data from port                                                                                   |
| `transmit_com_port(1,"QWERTY")` |  Yes | Sends string to port                                                                                          |
