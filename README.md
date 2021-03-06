# PrepServer
A tool to optimize virtual machines before provisioning.

## **Key features**
- Checks if there's a pending reboot before actually shutting down.
- Autologon and reboot-loop until the system no longer requires reboot, then runs shutdown.
- Place bat- and registry files in their respective directories to have them run automatically before shutdown.
- Deletes Windows Update files. (if Windows Update has finished, and if there's no pending reboot)
- Automatically resets RDS grace period timebomb.
- Clears DNS- and Group Policy cache.
- Automatically disables tasks and services.

## Usage
- Place **PrepServer.exe** on your VM.
- When you're ready to shutdown, run **PrepServer.exe**.

## Preview
PrepServer takes up the whole screen. **To exit, hit '_shift + alt + Q_'**.

![image](https://user-images.githubusercontent.com/93126880/138739598-35ec9090-ecd0-481d-96a2-112d0e3aaa6c.png)

![image](https://user-images.githubusercontent.com/93126880/138739641-4b23bc25-779b-46a3-8878-c61306f03bcf.png)

A pending reboot was detected, and a countdown runs before rebooting.
Again, use the hotkey to exit.

If Autologon info is found in 'Config.ini', this will be written to the registry and used to sign in automatically after reboot.

## Compile

### **IMPORTANT**

Copy ALL **au3** files from **_Includes_** and replace your existing ones, here:

C:\Program Files (x86)\AutoIt3\Include

Alternatively, you can point the script to Include from a different directory.

The **au3** files from this repo may break other scripts.

I'm considering changing this in the source code.

## Disclaimer
Please note that this is a work-in-progress.

You will find bloated code and unnecessary/unused code and comments.

## Credit
There are some clever people in the AutoIt community.
I've used a bunch of their code - and I got alot of inspiration from their work.
