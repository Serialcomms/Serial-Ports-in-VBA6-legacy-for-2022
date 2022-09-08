## VBA Serial Port routines for Legacy Microsoft Windows and Office editions
### New for 2022 - Windows XP onwards, pre-Office 2010 editions (32-Bit VBA6)

Getting Serial (COM) Ports working as intended in VBA can be surprisingly difficult in certain usage scenarios. 

New VBA routines here will help resolve the above in Excel, Word and Access (Windows PC versions only).

Functions are straightforward to use and coding style supports infrequent VBA users and developers.

Intended to help implement ad-hoc projects for serial data acquisition or transfer.

No plug-ins, DLLs, ActiveX, licences, payments or registrations are required.  

<details><summary>More Information</summary>
<p>
   
<details><summary>VBA Com Port function issues</summary>
<p>

The in-built VBA functions for COM Port data can suffer from the following issues :- 
   
1. Setting port parameters with the VBA open command may not work in some Windows versions e.g.

   `Open "COM1:9600,N,8,1" For Read Access As #1`       \
     _(command line workaround known, settings can revert after reboot)_

2. Attempting to read data when there is none waiting will cause VBA to hang with a 'not responding' message.  
  
   `Get #1, , Read_Data_Byte`  
   
   The new functions address both of these issues, and also where data transfers take longer than the 5-6 second VBA timeout.
   
</p>
</details>   
   
The legacy of serial comms means that many existing VBA solutions are now time-expired with links to defunct web sites etc.    

New functions here are a fresh start for 2022 and are based largely on Microsoft's Win32 API calls and documentation. 
   
Cloned from https://github.com/serialcomms/Serial-Ports-in-VBA-new-for-2022.git and modified for VBA6.

Minimal testing only on Windows XP Pro, SP3 with hardware vendor bundled version of Microsoft Word 2002.

Functions are straightforward to use and intended to help implement ad-hoc projects for serial data acquisition or transfer.



  
<details><summary>Debugging</summary>
<p>

* Debugging can be set on/off per port with results shown in the VBA immediate window. 

* Extensive debug functionality makes several modules quite verbose. 

* A far more compact version without debug is available in the No-Debug folder. 

</p>
</details>   


<details><summary>COM Ports</summary>
<p>

Multiple com ports are supported, including physical hardware ports and virtual software ports. 

All read and write functions are synchronous, in part because not all serial port types support overlapped operation.

</p>
</details>   
 
Performance on a modern PC is good, with software timing delays required to allow the relatively slow serial com ports to catch up. 

Reading, Writing and Waiting are 'timesliced' to ensure that VBA remains responsive during any extended data transfers or waiting times. 

<details><summary>Optional steps for Excel only</summary>
<p>

- Remove comment mark before `Option Private Module` to prevent function names appearing in cell formula drop-down lists. 
- Remove comment mark before `Application.Volatile` where indicated to refresh results when functions are used in cells and the worksheet is recalculated (e.g. with F9 key).

</p>
</details>

[Main user-defined functions](Functions/Function_List_All.md)

</p>
</details>   
