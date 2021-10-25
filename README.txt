Author: André Saga Lande


Developed as a tool to optimize virtual machines before provisioning.
Use in conjunction with other tools and scripts.


Key features
- Checks if there's a pending reboot before actually shutting down.
- Autologon and reboot-loop until the system no longer requires reboot, then runs shutdown.
- Easy to add features. Place bat- and registry files in their respective directories to have them run automatically before shutdown.
- Deletes Windows Update files. (if Windows Update has finished, and if there's no pending reboot)
- Automatically resets RDS grace period timebomb.
- Clears DNS cache and Group Policy cache.
- Automatically disables tasks and services.