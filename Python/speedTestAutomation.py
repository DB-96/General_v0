#!/usr/bin/env python
###################################################################################################
## This script runs with a cronjob on a Linux System
## Access/open the crontab editor/nano file using the below command;
## 'crontab' -e // or to list your existing jobs use 'crontab -l'
## The below command runs the wifi "speedtest" every 45 minutes and appends it to a .csv file for
## later processing
## 45 * * * * /usr/bin/python3 /home/db/Development/Git/Python_Natives/speedTestAutomation.py
###################################################################################################
import speedtest as st
import pandas as pd

import socket as soc

from datetime import datetime

def get_new_speeds():
    speed_test = st.Speedtest()
    speed_test.get_best_server()
    # Ping (ms)
    ping = speed_test.results.ping
    # Upload/Download
    download     = speed_test.download()
    upload       = speed_test.upload()
    download_mbs = round(download/(10**6),3)
    upload_mbs   = round(upload/(10**6),3)

    return(ping, download_mbs, upload_mbs)

def update_csv(internet_speeds):
    date_today =datetime.now().strftime("%d-%m-%Y %H:%M:%S ")
    csv_file_name   = "/home/db/Development/Git/Python/TestAutomation.csv"

    try:
        csv_dataset = pd.read_csv(csv_file_name, index_col="Date")
    except:
        csv_dataset = pd.DataFrame(
            list(),
            columns=["Ping (ms)","Download (Mb/s)","Upload (Mb/s)", "IP Address"]
        )
    # hostname   = soc.getfqdn() 
    # IPAddr     = soc.gethostbyname(soc.getfqdn())
    hostname   = soc.socket(soc.AF_INET, soc.SOCK_DGRAM)
    hostname.connect(("8.8.8.8", 80))
    IPAddr     = hostname.getsockname()[0]
    #  192.168.178.192 is typically the ethernet connection
 
    results_df = pd.DataFrame(
        [[internet_speeds[0], internet_speeds[1], internet_speeds[2], IPAddr]],
        columns=["Ping (ms)", "Download (Mb/s)","Upload (Mb/s)", "IP Address"],
        index = [date_today]
    )
    updated_df = csv_dataset.append(results_df)
    updated_df\
        .loc[~updated_df.index.duplicated(keep="last")]\
            .to_csv(csv_file_name, index_label = "Date")

new_speeds = get_new_speeds()
update_csv(new_speeds)
print(new_speeds)