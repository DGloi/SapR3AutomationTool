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
    pip install questionary
```




 ---
 
 ## How does that work ?


A picture is worth a thousand words:





---
## How can I record my own SAP Tcode ? 

The process to record your own process from sap into a vba code is quite straightforward:

* Open SAP Logon, and then select the SAP system to which you want to sign in.

![sap-login-760](https://user-images.githubusercontent.com/93589158/197603933-ac2a4bce-5858-4917-abba-f20afce73b0f.png)

* Select Customize Local Layout (Alt F12), and then select Script Recording and Playback.

![sap-easy-access-system](https://user-images.githubusercontent.com/93589158/197608516-1f55dc28-2550-455c-9d3c-25acdb701fde.png)

*Select More.

Under Save To, provide the path and file name where you want to store the captured user interactions.

![saving-recording-file](https://user-images.githubusercontent.com/93589158/197609231-56c10887-9360-4953-8e4d-534c4268b031.png)
