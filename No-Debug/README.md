# Non-Debug Version
This is the 'lightweight' version of SERIAL_PORTS_VBA with all debug and show functionality removed.

All Main VBA port functions remain functionally equivalent. 

Port debug function remains for compatibility and always returns false.

Cloned from https://github.com/serialcomms/Serial-Ports-in-VBA-new-for-2022.git and modified for VBA6 usage

* All `LongPtr` references changed to `Long`

* All `PtrSafe` references in function declarations removed


Suggestion is to use full debug version initially for testing, and then change to this version for regular use if appropriate. 
