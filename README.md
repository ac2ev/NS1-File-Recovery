All the way from 2007
	

How to repair a Corrupted NS1 file
Assumptions:  This page assumes you know how to download and extract a file to your own system.  If you don't know how to do that, you probably shouldn't be attempting this procedure.


Method 1 AP COUNT Manipulation – Here for historical purposes Use Method 2 below

NS1 Header Format:
	

4E 65 74 53
	

0c 00 00 00
	

1C 00 00 00

 
	

NetS dwSignature
	

dwFileversion
	

AP Count (little endian first)

 

·         Step 1:  Download and extract the File from the "Required file Download" link Above.

·         Step 2:  Run XVi32.exe from the extracted directory.

·         Step 3:  Make a backup copy of all your NS1 files you will be attempting to repair, in case something goes horribly wrong.

·         Step 4:  Open the corrupt NS1 file in Xvi32.

Select the Beginning of the AP Count (Goto Address)

Byte 8 selected

Decode the value to see what the APCount is currently set at

Encode a new value

 

·         Step 5:  Try to open the file in Netstumbler

        If the file opens proceed to Step 7
        If the file fails to open the then you need to use Beakmyn’s Ns1 File Recovery Program
    Step 6:  Set the AP count to the approximate number of AP's before the file crashed.
        If the file opens, increase the AP count by 20 and open until the file fails to open in Netstumbler, proceed to Step 8.
        If the file fails to open, decrease the AP count by 20 until the file opens in Netstumbler, proceed to Step 8.
    Step 7:  Increase or Decrease the AP count by 1 until the file fails to open then back off the AP count by 1.
    Step 8:  You have Recovered all AP's that can be recovered using this method. For full recovery use Beakmyn’s Ns1 File Recovery Program


 

