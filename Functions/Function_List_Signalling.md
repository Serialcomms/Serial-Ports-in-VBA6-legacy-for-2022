### VBA functions to check and set port signals.

#### Input signals


First parameter (1) is a valid[^1] COM Port number on host PC

| VBA Function                  | Signal | Description                                  |
| ------------------------------|------- | ---------------------------------------------|
| `device_ready(1)`             | DSR    | Data Set Ready                               |
| `device_calling(1)`           | RI     | Ring Indicate                                |
| `carrier_detect(1)`           | CD     | Carrier Detect / Receive Line Signal Detect  |
| `clear_to_send(1)`            | CTS    | Clear To Send                                |  
 
 * Functions return True if port valid, started and input signal active.

#### Output signals[^2]

First parameter (1) is a valid[^1] COM Port number on host PC

| VBA Function                  | Signal | Action | Description                          |
| ------------------------------|------- | -------|--------------------------------------|
| `request_to_send(1,0)`        | RTS    |  Clear | Request To Send                      |
| `request_to_send(1,1)`        | RTS    |  Set   | Request To Send                      |
| `signal_com_port(1,1)`        | XOFF   |  Set   | Flow Control Off                     |
| `signal_com_port(1,2)`        | XON    |  Set   | Flow Control On                      |
| `signal_com_port(1,3)`        | RTS    |  Set   | Request To Send                      |
| `signal_com_port(1,4)`        | RTS    |  Clear | Request To Send                      |
| `signal_com_port(1,5)`        | DTR    |  Set   | Data Terminal Ready                  |
| `signal_com_port(1,6)`        | DTR    |  Clear | Data Terminal Ready                  |
| `signal_com_port(1,8)`        | BREAK  |  Set   | Request To Send                      |
| `signal_com_port(1,9)`        | BREAK  |  Clear | Request To Send                      |

 * Functions return True if port valid, started and Windows EscapeCommFunction[^2] returns True 

[^1]: valid
[^2]: see - [Escape Comm Function signal values](https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-escapecommfunction)
