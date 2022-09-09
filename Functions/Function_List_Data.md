### VBA functions to send, receive and check for data

First parameter (1) is a Valid[^1] and Started COM Port number on host PC.

| VBA Function                    |  TS  | Description                                                                                                   |
| :-------------------------------|:----:| :-------------------------------------------------------------------------------------------------------------|
| `check_com_port(1)`             | No   | Returns number of input characters waiting to be read (no wait). Return value -1 indicates error.             |
| `wait_com_port(1)`              | Yes  | Wait for up to 333mS (default) before timing out. Returns `True` if receive data waiting.                     |
| `wait_com_port(1,500)`          | Yes  | As above, specify maximum wait time (500) in milliseconds.                                                    |  
| `get_com_port(1)`               | No   | Receives a single-character string.                                                                           |
| `put_com_port(1,"A")`           | No   | Sends a single-character string.                                                                              |
| `read_com_port(1,20)`           | No   | Reads up to specified number (20) of characters. No delay[^2]                                                 |
| `send_com_port(1,V)`            | Yes  | Sends variable V. Function converts V to string and calls `transmit_com_port`.                                |
| `receive_com_port(1)`           | Yes  | Receives all data from port                                                                                   |
| `transmit_com_port(1,"QWERTY")` | Yes  | Sends supplied string to port                                                                                 |

* Functions shown as TS=Yes are timesliced to avoid VBA hanging with a 'not responding' message.
[^1]:  Valid Minimum and Maximum ports are defined in declarations section at the top of the module. 
[^2]:  Maximum characters read is approximately com port baud rate / 10

