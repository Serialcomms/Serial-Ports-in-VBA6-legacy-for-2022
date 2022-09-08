Main user-defined port Signalling functions are as follows. First parameter (1) is a valid COM Port number on host.

| VBA Function                         | Description                                                                                                   |
| ------------------------------------ | --------------------------------------------------------------------------------------------------------------|
| `device_ready(1)`                    | Returns `True` if port started and Data Set Ready (DSR) input signal active.                                  |
| `device_calling(1)`                  | Returns `True` if port started and Ring Indicate (RI) input signal active.                                    |
| `carrier_detect(1)`                  | Returns `True` if port started and Carrier Detect (CD) input signal active.                                   |
| `clear_to_send(1)`                   | Returns `True` if port started and Clear To Send (CTS) input signal active.                                   |
| `request_to_send(1,1)`               | Sets Request To Send (RTS) output signal. Returns `True` if port started and RTS set                          |
| `request_to_send(1,0)`               | Clears Request To Send (RTS) output signal. Returns `True` if port started and RTS cleared                    |
| `signal_com_port(1,signal)`          | Set/clear BREAK, DTR, RTS output signals, see - [Escape Comm Function signal values](https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-escapecommfunction)
