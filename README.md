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
