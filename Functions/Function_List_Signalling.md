### VBA functions to check and set port signals.

#### Input signals


First parameter (1) is a valid[^1] COM Port number on host PC

| VBA Function                  | Signal | Description                                  | DB9 Pin | DB25 Pin | 
| ------------------------------|------- | ---------------------------------------------|:-------:|:--------:|
| `device_ready(1)`             | DSR    | Data Set Ready                               |    6    |    6     |
| `device_calling(1)`           | RI     | Ring Indicate                                |    9    |    22    |
| `carrier_detect(1)`           | CD     | Carrier Detect / Receive Line Signal Detect  |    1    |    8     |
| `clear_to_send(1)`            | CTS    | Clear To Send                                |    8    |    5     | 
 
 * Functions return True if port valid, started, input signal active and Windows GetCommModemStatus returns True

#### Output signals[^2]

First parameter (1) is a valid[^1] COM Port number on host PC

| VBA Function                  | Signal | Action | Description                          | DB9 Pin | DB25 Pin | 
| ------------------------------|------- | -------|--------------------------------------|:-------:|:--------:|
| `request_to_send(1,0)`        | RTS    |  Clear | Request To Send                      |    7    |    4     |
| `request_to_send(1,1)`        | RTS    |  Send  | Request To Send                      |    7    |    4     |
| `signal_com_port(1,1)`        | XOFF   |  Send  | Flow Control Off                     |    -    |    -     |
| `signal_com_port(1,2)`        | XON    |  Send  | Flow Control On                      |    -    |    -     |
| `signal_com_port(1,3)`        | RTS    |  Send  | Request To Send                      |    7    |    4     |
| `signal_com_port(1,4)`        | RTS    |  Clear | Request To Send                      |    7    |    4     |
| `signal_com_port(1,5)`        | DTR    |  Send  | Data Terminal Ready                  |    4    |    20    |
| `signal_com_port(1,6)`        | DTR    |  Clear | Data Terminal Ready                  |    4    |    20    |
| `signal_com_port(1,8)`        | BREAK  |  Send  | Line Break Condition                 |    3    |    3     |
| `signal_com_port(1,9)`        | BREAK  |  Clear | Line Break Condition                 |    3    |    3     |

 * Functions return True if port valid, started and Windows EscapeCommFunction[^2] returns True 

[^1]: Valid Minimum and Maximum port numbers should be defined in declarations section at the start of the module.
[^2]: see - [Escape Comm Function signal values](https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-escapecommfunction)
