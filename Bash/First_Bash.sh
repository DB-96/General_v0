#!/bin/bash
touch DateTimeFile.txt
while true; do
    date | cat >> DateTimeFile.txt
    echo "Waiting"
    sleep 5  ## sleep always defined in seconds
done
## use the chmod command to allow this file executable permissions (i.e. it can run as an executable and write to a text file)
## run "./First_Bash.sh"
## Ctrl + C will terminate the shell script in the terminal