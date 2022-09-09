### VBA functions to check and set port signals.


First parameter (1) is a valid[^1] COM Port number on host.

| VBA Function                  | Signal | Description                                                                                                   |
| ------------------------------|------- | --------------------------------------------------------------------------------------------------------------|
| `device_ready(1)`             | DSR    | Data Set Ready                               |
| `device_calling(1)`           | RI     | Ring Indicate                                   |
| `carrier_detect(1)`           | CD     | Carrier Detect                                  |
| `clear_to_send(1)`            | CTS    | Clear To Send                                |
| `request_to_send(1,1)`        | RTS    | Sets Request To Send Returns `True` if port started and RTS set                          |
| `request_to_send(1,0)`        | RTS    | Clears Request To Send Returns `True` if port started and RTS cleared                    |
| `signal_com_port(1,signal)`   |        | Set/clear BREAK, DTR, RTS output signals, see - [Escape Comm Function signal values](https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-escapecommfunction)

[^1]: valid
