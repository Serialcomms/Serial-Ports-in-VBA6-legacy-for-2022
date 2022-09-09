### VBA functions to check and set port signals.

#### Input signals


First parameter (1) is a valid[^1] and started COM Port number on host PC

| VBA Function                  | Signal | Description                                  |
| ------------------------------|------- | ---------------------------------------------|
| `device_ready(1)`             | DSR    | Data Set Ready                               |
| `device_calling(1)`           | RI     | Ring Indicate                                |
| `carrier_detect(1)`           | CD     | Carrier Detect / Receive Line Signal Detect  |
| `clear_to_send(1)`            | CTS    | Clear To Send                                |  
 
 * Functions return True if port valid, started and input signal active.

| VBA Function                  | Signal | Description                             |
 |------------------------------|------- | ----------------------------------------|
| `request_to_send(1,1)`        | RTS    | Sets Request To Send                    |
| `request_to_send(1,0)`        | RTS    | Clears Request To Send                  |
| `signal_com_port(1,signal)`   |        | Set/clear BREAK, DTR, RTS output signals, [^2]:


#### Output signals

[^1]: valid
[^2]: see - [Escape Comm Function signal values](https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-escapecommfunction)
