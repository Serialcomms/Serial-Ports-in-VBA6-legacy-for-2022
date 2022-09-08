Main user-defined Data functions are as follows. First parameter (1) is a valid COM Port number on host.

| VBA Function                         | Description                                                                                                   |
| ------------------------------------ | --------------------------------------------------------------------------------------------------------------|
| `check_com_port(1)`                  | Returns number of input characters waiting to be read (no delay). Return value -1 indicates error.            |
| `wait_com_port(1)`                   | Wait for up to 333mS (default) before timing out. Returns `True` if receive data waiting.                     |
| `wait_com_port(1,500)`               | As above, can optionally specify wait time (500) in milliseconds. Timesliced to avoid VBA hanging.            |  
| `get_com_port(1)`                    | Receives a single character string from a started com port.                                                   |
| `put_com_port(1,"A")`                | Sends a single character to a started com port.                                                               |
| `read_com_port(1,20)`                | Reads up to specified number (20) of characters. No delay, max characters = approx 1 second timeslice.        |
| `send_com_port(1,V)`                 | Sends variable V. Function converts V to String and calls `transmit_com_port`.                                |
| `receive_com_port(1)`                | Receives all data from port, timesliced for low port speeds and/or large data transfers.                      |
| `transmit_com_port(1,"QWERTY")`      | Sends string to port, in timeslices of approx 1 second to avoid VBA 'not responding'                          |
