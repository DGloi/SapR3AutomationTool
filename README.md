![r3-1601913976](https://user-images.githubusercontent.com/93589158/197033907-db6a521c-a1f5-4c7e-a7a4-7d1ead778996.png)
# Sap-R3-Automation-tool
 SAP R3 is still used by numerous companies and many people still use it everyday at their work, sometimes running the same reports all over again everyday.
 This is where this solution comes to play. 
 
This repositeries provides a framework to be applied to any situation described above. It cannot be simply forked or copied to make it work on your situation as you will be requiered to do some modifications and record your own process in SAP.


## Technologies

This project leverages microsoft excel, VBA scripting language and python 3.6 with the following packages:

* [pywin32](https://pypi.org/project/pywin32/)



---

## Installation Guide
Before running the application, first install the following dependencies:
```python
    pip install pywin32
```




 ---
 
 ## How does that work ?

SAP R3 allows you to record your own reports in VBA by default so you can simply launch the VBA script later to let your machine fetch the data on its own.
(Recording process below).
So the idea is to have programs to log us in SAP R3 in the right SAP environment, launch and save the needed reports , clean the data and begin analysis.

Still a bit confused ?

A picture is worth a thousand words:



Basically, the first script Sap_login.py will allow you to connect automatically to your SAP interface. Then vba_macros_launcher.py will be able to launch every macro recorded from SAP and stored into the excel file VBA_Macro_4_sapR3.xlsm

---
## How can I record my own SAP Tcode ? 

The process to record your own process from sap into a vba code is quite straightforward:

* Open SAP Logon, and then select the SAP system to which you want to sign in.

![sap-login-760](https://user-images.githubusercontent.com/93589158/197603933-ac2a4bce-5858-4917-abba-f20afce73b0f.png)

* Select Customize Local Layout (Alt F12), and then select Script Recording and Playback.

![sap-easy-access-system](https://user-images.githubusercontent.com/93589158/197608516-1f55dc28-2550-455c-9d3c-25acdb701fde.png)

* Select More.

Under Save To, provide the path and file name where you want to store the captured user interactions.

![saving-recording-file](https://user-images.githubusercontent.com/93589158/197609231-56c10887-9360-4953-8e4d-534c4268b031.png)
