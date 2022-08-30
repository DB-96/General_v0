#!/bin/bash
# touch WiFiSpeed_DateTime.txt
# while true; do

echo "$(date), $(ip r | grep default), $(speedtest --simple)" >> log.csv

#    date  | cat >> WiFiSpeed_DateTime.txt
#    ip r | grep default | cat >> WiFiSpeed_DateTime.txt
#    speedtest --simple  | cat >> WiFiSpeed_DateTime.txt

# sleep 5  ## sleep always defined in seconds
# done
## use the chmod command to allow this file executable permissions (i.e. it can run as an executable and write to a text file)
## run "./First_Bash.sh"
## Ctrl + C will terminate the shell script in the terminal