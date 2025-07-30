import os
os.system("") # enables ansi escape code processing for cmd
import requests # for importing functions to fetch pi data for a tag
import getpass # for accessing user credentials for pi data
import urllib3 # to remove ssl cert bypass warnings
from datetime import datetime # for timestamp
import sys
import pandas as pd # for export to csv
from openpyxl import load_workbook # for writing to excel
import matplotlib.pyplot as plt
import mplcursors

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning) # ignore warnings

logo = r'''
      _    _ ______  _____ _____ 
     | |  | |  ____|/ ____/ ____|
     | |__| | |__  | (___| (___  
     |  __  |  __|  \___ \\___ \ 
     | |  | | |____ ____) |___) |
     |_|  |_|______|_____/_____/
    '''

print(f"\033[92m{logo}\033[0m")

print("\033[92mWelcome to Hess SoR Cooler Analysis Tool\033[0m\n")

flag = True

while flag is True:
    username = input("Username: \n")
    password = getpass.getpass("Password: \n")

    logintest = requests.get("https://ndpi.vision.ihess.com/piwebapi/", auth=(username, password), verify=False).json()
    if "Authorization has been denied for this request." in logintest.values(): # verifies if user is logged in correctly
        print("Incorrect Login. Try Again\n")
    else:
        flag = False

print("Loading...\n")

station = input("Which compressor station, HGF | BBCS2 | BWCS | BBCS Comp 4 ? \n").strip().upper() # user types station type

# HGF coolers code
if station == "HGF" or station == '0':
    print("You have selected HGF\n")
    cpHGF = 0.54375 # Cp of inlet gas [BTU/lb/°F]
    mmHGF = 26 # molecular mass of inlet gas [lb/lb-mol]
    
    while True:
        cooler = input("Which cooler, Inlet Gas | Flash Gas | Refrig ? \n").strip().lower() # user types cooler type for hgf
    
        if cooler == 'inlet gas' or cooler == '0': # if hgf inlet gas coolers selected
            print("You have selected HGF Inlet Gas Coolers\n")
            
            # --- COOLER 1 --- 
            volHGF = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQI7wDAATkRQ" \
                "SURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpIQVdLRVlFLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5" \
                "MRVRSRUNJUDEuQ0FMQ1VMQVRFRkxPV1lEWV80QU0uRlFJVDc1MjEwMQ/value",
                auth=(username, password), verify=False).json() # gets pi data (flowrate and timestamp) for Cooler 1
            if not isinstance(volHGF["Value"], (int, float)): # if value has bad input
                print("\033[91mError: Bad transmitter input for Cooler 1 flow...\nDuty marked as 0 BTU/hr...\033[0m")  
                mHGF = 0
            else:
                mHGF = volHGF["Value"] * 1/397.3 * mmHGF / 24 # gets mass flow rate [1000-lb/hr]

            # -- INTERCOOLER 1 --

            # gaining values for q computation
            input("Verify that the operator rounds data is uploaded to this folder with the name 'InletRounds.xlsx' and press ENTER to continue...\n")
            print("\033[94mNow entering for Cooler 1...\033[0m\n")
            T_in_list = pd.read_excel("InletRounds.xlsx")[['Label', 'Response']]
            status = (T_in_list[T_in_list['Label'] == 'COMPRESSOR STATUS'])['Response'].tolist()
            T_in_list['Response'] = pd.to_numeric((T_in_list['Response'].astype(str).str.replace(' degF', '', regex=False)), errors='coerce')
            if status[0].strip().upper() == 'OOS' or status[0].strip().upper() == 'STANDBY':
                print("\033[91mError: Comp 1 OOS...\nUnable to compute duties. Exiting program...\033[0m")  
                sys.exit()
            else:
                T_in = (T_in_list[T_in_list['Label'] == '1ST STAGE #1 NORTH OUTLET TEMP'].iloc[0])['Response']
                T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQbqgBAATkRQSURBQ09MTF" \
                    "xVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpIQVdLRVlFLk" \
                    "dBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDEuRklSU1RTVEFHRUlOVEVSQ09PTEVSVEVNUC5USVQ3NTIxMTE/value",
                    auth=(username, password), verify=False).json() # gets pi data for outlet temp for Cooler 1, IC-1
                if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                    print("\033[91mError: Bad transmitter input for Cooler 1 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
                    q_C1_IC1 = 0
                else:
                    q_C1_IC1 = mHGF * cpHGF * (T_in - T_out["Value"]) # heat rejection (1000-BTU/hr)

                # -- INTERCOOLER 2 --

                # gaining values for q computation
                T_in = (T_in_list[T_in_list['Label'] == '2ND STAGE NORTH OUTLET TEMP'].iloc[0])['Response']
                T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQclsBAATkRQSURBQ0" \
                    "9MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpIQVdLRVlFLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUD" \
                    "EuU0VDT05EU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUNzUyMTEy/value",
                    auth=(username, password), verify=False).json() # gets pi data for outlet temp for Cooler 1, IC-2
                if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                    print("\033[91mError: Bad transmitter input for Cooler 1 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
                    q_C1_IC2 = 0
                else:
                    q_C1_IC2 = mHGF * cpHGF * (T_in - T_out["Value"]) # heat rejection (1000-BTU/hr)

                # -- INTERCOOLER 3 --

                # gaining values for q computation
                T_in = (T_in_list[T_in_list['Label'] == '3RD STAGE OUTLET TEMP'].iloc[0])['Response']
                T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQdlsBAATkRQSURB" \
                    "Q09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpIQVdLRVlFLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJU" \
                    "DEuVEhJUkRTVEFHRUlOVEVSQ09PTEVSVEVNUC5USVQ3NTIxMTM/value",
                    auth=(username, password), verify=False).json() # gets pi data for outlet temp for Cooler 1, IC-3
                if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                    print("\033[91mError: Bad transmitter input for Cooler 1 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
                    q_C1_IC3 = 0
                else:
                    q_C1_IC3 = mHGF * cpHGF * (T_in - T_out["Value"]) # heat rejection (1000-BTU/hr)

                # -- INTERCOOLER 4 --

                # gaining values for q computation
                T_in = (T_in_list[T_in_list['Label'] == '4TH STAGE OUTLET TEMP'].iloc[0])['Response']
                T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQcagBAATkRQ" \
                    "SURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpIQVdLRVlFLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5" \
                    "MRVRSRUNJUDEuRk9VUlRIU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUNzUyMTE0/value",
                    auth=(username, password), verify=False).json() # gets pi data for outlet temp for Cooler 1, IC-4
                if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                    print("\033[91mError: Bad transmitter input for Cooler 1 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
                    q_C1_IC4 = 0
                else:
                    q_C1_IC4 = mHGF * cpHGF * (T_in - T_out["Value"]) # heat rejection (1000-BTU/hr)

                # -- EJW

                T_in = (T_in_list[T_in_list['Label'] == 'ENGINE: JACKET WATER OUTLET TEMP'].iloc[0])['Response']
                T_out = (T_in_list[T_in_list['Label'] == 'ENGINE: JACKET WATER INLET TEMP'].iloc[0])['Response']
                q_C1_EJW = ((575 * 192.5 * 1/24) * (1.02 * 60.793))  * (0.643) * (T_in - T_out) / 1000 # heat rejection (1000-BTU/hr)

            # --- COOLER 2 --- 
            print("\n\033[94mNow entering for Cooler 2...\033[0m\n")
            if status[1].strip().upper() == 'OOS' or status[1].strip().upper() == 'STANDBY':
                print("\033[91mError: Comp 2 OOS...\nUnable to compute duties. Exiting program...\033[0m")  
                sys.exit()
            else:
                volHGF = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQJbwDAATkR" \
                    "QSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpIQVdLRVlFLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU" \
                    "5MRVRSRUNJUDIuQ0FMQ1VMQVRFRkxPV1lEWV80QU0uRlFJVDc1MjIwMQ/value",
                    auth=(username, password), verify=False).json() # gets pi data (flowrate and timestamp) for Cooler 2
                if not isinstance(volHGF["Value"], (int, float)): # if value has bad input
                    print("\033[91mError: Bad transmitter input for Cooler 2 flow...\nDuty marked as 0 BTU/hr...\033[0m")  
                    mHGF = 0
                else:
                    mHGF = volHGF["Value"] * 1/397.3 * mmHGF / 24 # gets mass flow rate [1000-lb/hr]

                # -- INTERCOOLER 1 --

                # gaining values for q computation 
                T_in = (T_in_list[T_in_list['Label'] == '1ST STAGE #1 NORTH OUTLET TEMP'].iloc[1])['Response']
                T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQeKgBAATkRQ" \
                    "SURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpIQVdLRVlFLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5" \
                    "MRVRSRUNJUDMuRklSU1RTVEFHRUlOVEVSQ09PTEVSVEVNUC5USVQ3NTIzMTE/value",
                    auth=(username, password), verify=False).json() # gets pi data for outlet temp for Cooler 2, IC-1
                if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                    print("\033[91mError: Bad transmitter input for Cooler 2 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
                    q_C2_IC1 = 0
                else:
                    q_C2_IC1 = mHGF * cpHGF * (T_in - T_out["Value"]) # heat rejection (1000-BTU/hr)

                # -- INTERCOOLER 2 --

                # gaining values for q computation
                T_in = (T_in_list[T_in_list['Label'] == '2ND STAGE NORTH OUTLET TEMP'].iloc[1])['Response']
                T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQeagBAATkRQ" \
                    "SURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpIQVdLRVlFLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5" \
                    "MRVRSRUNJUDMuU0VDT05EU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUNzUyMzEy/value",
                    auth=(username, password), verify=False).json() # gets pi data for outlet temp for Cooler 2, IC-2
                if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                    print("\033[91mError: Bad transmitter input for Cooler 2 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
                    q_C2_IC2 = 0
                else:
                    q_C2_IC2 = mHGF * cpHGF * (T_in - T_out["Value"]) # heat rejection (1000-BTU/hr)

                # -- INTERCOOLER 3 --

                # gaining values for q computation
                T_in = (T_in_list[T_in_list['Label'] == '3RD STAGE OUTLET TEMP'].iloc[1])['Response']
                T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQeqgBAATkRQS" \
                    "URBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpIQVdLRVlFLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5M" \
                    "RVRSRUNJUDMuVEhJUkRTVEFHRUlOVEVSQ09PTEVSVEVNUC5USVQ3NTIzMTM/value",
                    auth=(username, password), verify=False).json() # gets pi data for outlet temp for Cooler 2, IC-3
                if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                    print("\033[91mError: Bad transmitter input for Cooler 2 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
                    q_C2_IC3 = 0
                else:
                    q_C2_IC3 = mHGF * cpHGF * (T_in - T_out["Value"]) # heat rejection (1000-BTU/hr)

                # -- INTERCOOLER 4 --

                # gaining values for q computation
                T_in = (T_in_list[T_in_list['Label'] == '4TH STAGE OUTLET TEMP'].iloc[1])['Response']
                T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQe6gBAATkRQ" \
                    "SURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpIQVdLRVlFLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU" \
                    "5MRVRSRUNJUDMuRk9VUlRIU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUNzUyMzE0/value",
                    auth=(username, password), verify=False).json() # gets pi data for outlet temp for Cooler 2, IC-4
                if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                    print("\033[91mError: Bad transmitter input for Cooler 2 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
                    q_C2_IC4 = 0
                else:
                    q_C2_IC4 = mHGF * cpHGF * (T_in - T_out["Value"]) # heat rejection (1000-BTU/hr)

                # -- EJW

                T_in = (T_in_list[T_in_list['Label'] == 'ENGINE: JACKET WATER OUTLET TEMP'].iloc[1])['Response']
                T_out = (T_in_list[T_in_list['Label'] == 'ENGINE: JACKET WATER INLET TEMP'].iloc[1])['Response']
                q_C2_EJW = ((575 * 192.5 * 1/24) * (1.02 * 60.793))  * (0.643) * (T_in - T_out) / 1000 # heat rejection (1000-BTU/hr)

            # --- COOLER 3 --- 

            # for constants and flows
            print("\n\033[94mNow entering for Cooler 3...\033[0m\n")
            if status[2].strip().upper() == 'OOS' or status[2].strip().upper() == 'STANDBY':
                print("\033[91mError: Comp 3 OOS...\nUnable to compute duties. Exiting program...\033[0m")  
                sys.exit()
            else:
                volHGF = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQJ7wDAATkR" \
                    "QSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpIQVdLRVlFLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU" \
                    "5MRVRSRUNJUDMuQ0FMQ1VMQVRFRkxPV1lEWV80QU0uRlFJVDc1MjMwMQ/value",
                    auth=(username, password), verify=False).json() # gets pi data (flowrate and timestamp) for Cooler 3
                if not isinstance(volHGF["Value"], (int, float)): # if value has bad input
                    print("\033[91mError: Bad transmitter input for Cooler 1 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
                    mHGF = 0
                else:
                    mHGF = volHGF["Value"] * 1/397.3 * mmHGF / 24 # gets mass flow rate [1000-lb/hr]

                # -- INTERCOOLER 1 --
                # gaining values for q computation
                T_in = (T_in_list[T_in_list['Label'] == '1ST STAGE #1 SOUTH OUTLET TEMP'].iloc[2])['Response']
                T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQfagBAATkRQ" \
                    "SURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpIQVdLRVlFLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5" \
                    "MRVRSRUNJUDQuRklSU1RTVEFHRUlOVEVSQ09PTEVSVEVNUC5USVQ3NTI0MTE/value",
                    auth=(username, password), verify=False).json() # gets pi data for outlet temp for Cooler 3, IC-1
                if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                    print("\033[91mError: Bad transmitter input for Cooler 3 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
                    q_C3_IC1 = 0
                else:
                    q_C3_IC1 = mHGF * cpHGF * (T_in - T_out["Value"]) # heat rejection (1000-BTU/hr)

                # -- INTERCOOLER 2 --

                # gaining values for q computation
                T_in = (T_in_list[T_in_list['Label'] == '2ND STAGE SOUTH OUTLET TEMP'].iloc[2])['Response']
                T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQfqgBAATkRQ" \
                    "SURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpIQVdLRVlFLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uS" \
                    "U5MRVRSRUNJUDQuU0VDT05EU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUNzUyNDEy/value",
                    auth=(username, password), verify=False).json() # gets pi data for outlet temp for Cooler 3, IC-2
                if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                    print("\033[91mError: Bad transmitter input for Cooler 3 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
                    q_C3_IC2 = 0
                else:
                    q_C3_IC2 = mHGF * cpHGF * (T_in - T_out["Value"]) # heat rejection (1000-BTU/hr)

                # -- INTERCOOLER 3 --

                # gaining values for q computation
                T_in = (T_in_list[T_in_list['Label'] == '3RD STAGE OUTLET TEMP'].iloc[2])['Response']
                T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQf6gBAATkR" \
                    "QSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpIQVdLRVlFLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04u" \
                    "SU5MRVRSRUNJUDQuVEhJUkRTVEFHRUlOVEVSQ09PTEVSVEVNUC5USVQ3NTI0MTM/value",
                    auth=(username, password), verify=False).json() # gets pi data for outlet temp for Cooler 3, IC-3
                if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                    print("\033[91mError: Bad transmitter input for Cooler 3 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
                    q_C3_IC3 = 0
                else:
                    q_C3_IC3 = mHGF * cpHGF * (T_in - T_out["Value"]) # heat rejection (1000-BTU/hr)

                # -- INTERCOOLER 4 --

                # gaining values for q computation
                T_in = (T_in_list[T_in_list['Label'] == '4TH STAGE OUTLET TEMP'].iloc[2])['Response']
                T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQgKgBAATkRQS" \
                    "URBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpIQVdLRVlFLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uS" \
                    "U5MRVRSRUNJUDQuRk9VUlRIU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUNzUyNDE0/value",
                    auth=(username, password), verify=False).json() # gets pi data for outlet temp for Cooler 3, IC-4
                if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                    print("\033[91mError: Bad transmitter input for Cooler 3 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
                    q_C3_IC4 = 0
                else:
                    q_C3_IC4 = mHGF * cpHGF * (T_in - T_out["Value"]) # heat rejection (1000-BTU/hr)

                # -- EJW

                T_in = (T_in_list[T_in_list['Label'] == 'ENGINE: JACKET WATER OUTLET TEMP'].iloc[2])['Response']
                T_out = (T_in_list[T_in_list['Label'] == 'ENGINE: JACKET WATER INLET TEMP'].iloc[2])['Response']
                q_C3_EJW = ((575 * 192.5 * 1/24) * (1.02 * 60.793))  * (0.643) * (T_in - T_out) / 1000 # heat rejection (1000-BTU/hr)

            # --- COOLER 4 --- 

            # for constants and flows
            print("\n\033[94mNow entering for Cooler 4...\033[0m\n")
            if status[3].strip().upper() == 'OOS' or status[3].strip().upper() == 'STANDBY':
                print("\033[91mError: Comp 4 OOS...\nUnable to compute duties. Exiting program...\033[0m")  
                sys.exit()
            else:
                volHGF = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQKbwDAATk" \
                    "RQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpIQVdLRVlFLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04" \
                    "uSU5MRVRSRUNJUDQuQ0FMQ1VMQVRFRkxPV1lEWV80QU0uRlFJVDc1MjQwMQ/value",
                    auth=(username, password), verify=False).json() # gets pi data (flowrate and timestamp) for Cooler 4
                if not isinstance(volHGF["Value"], (int, float)): # if value has bad input
                    print("\033[91mError: Bad transmitter input for Cooler 4 flow...\nDuty marked as 0 BTU/hr...\033[0m")  
                    mHGF = 0
                else:
                    mHGF = volHGF["Value"] * 1/397.3 * mmHGF / 24 # gets mass flow rate [1000-lb/hr]

                # -- INTERCOOLER 1 --
                # gaining values for q computation
                T_in = (T_in_list[T_in_list['Label'] == '1ST STAGE #1 NORTH OUTLET TEMP'].iloc[3])['Response']
                T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQfagBAATkRQ" \
                    "SURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpIQVdLRVlFLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5" \
                    "MRVRSRUNJUDQuRklSU1RTVEFHRUlOVEVSQ09PTEVSVEVNUC5USVQ3NTI0MTE/value",
                    auth=(username, password), verify=False).json() # gets pi data for outlet temp for Cooler 4, IC-1
                if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                    print("\033[91mError: Bad transmitter input for Cooler 4 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
                    q_C4_IC1 = 0
                else:
                    q_C4_IC1 = mHGF * cpHGF * (T_in - T_out["Value"]) # heat rejection (1000-BTU/hr)

                # -- INTERCOOLER 2 --

                # gaining values for q computation
                T_in = (T_in_list[T_in_list['Label'] == '2ND STAGE OUTLET TEMP'].iloc[0])['Response']
                T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQfqgBAATkRQ" \
                    "SURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpIQVdLRVlFLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uS" \
                    "U5MRVRSRUNJUDQuU0VDT05EU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUNzUyNDEy/value",
                    auth=(username, password), verify=False).json() # gets pi data for outlet temp for Cooler 4, IC-2
                if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                    print("\033[91mError: Bad transmitter input for Cooler 4 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
                    q_C4_IC2 = 0
                else:
                    q_C4_IC2 = mHGF * cpHGF * (T_in - T_out["Value"]) # heat rejection (1000-BTU/hr)

                # -- INTERCOOLER 3 --

                # gaining values for q computation
                T_in = (T_in_list[T_in_list['Label'] == '3RD STAGE OUTLET TEMP'].iloc[3])['Response']
                T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQf6gBAATkR" \
                    "QSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpIQVdLRVlFLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04u" \
                    "SU5MRVRSRUNJUDQuVEhJUkRTVEFHRUlOVEVSQ09PTEVSVEVNUC5USVQ3NTI0MTM/value",
                    auth=(username, password), verify=False).json() # gets pi data for outlet temp for Cooler 4, IC-3
                if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                    print("\033[91mError: Bad transmitter input for Cooler 4 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
                    q_C4_IC3 = 0
                else:
                    q_C4_IC3 = mHGF * cpHGF * (T_in - T_out["Value"]) # heat rejection (1000-BTU/hr)

                # -- INTERCOOLER 4 --

                # gaining values for q computation
                T_in = (T_in_list[T_in_list['Label'] == '4TH STAGE OUTLET TEMP'].iloc[3])['Response']
                T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQgKgBAATkRQS" \
                    "URBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpIQVdLRVlFLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uS" \
                    "U5MRVRSRUNJUDQuRk9VUlRIU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUNzUyNDE0/value",
                    auth=(username, password), verify=False).json() # gets pi data for outlet temp for Cooler 4, IC-4
                if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                    print("\033[91mError: Bad transmitter input for Cooler 4 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
                    q_C4_IC4 = 0
                else:
                    q_C4_IC4 = mHGF * cpHGF * (T_in - T_out["Value"]) # heat rejection (1000-BTU/hr)

                # -- EJW

                T_in = (T_in_list[T_in_list['Label'] == 'ENGINE: JACKET WATER OUTLET TEMP'].iloc[3])['Response']
                T_out = (T_in_list[T_in_list['Label'] == 'ENGINE: JACKET WATER INLET TEMP'].iloc[3])['Response']
                q_C4_EJW = ((575 * 192.5 * 1/24) * (1.02 * 60.793))  * (0.643) * (T_in - T_out) / 1000 # heat rejection (1000-BTU/hr)

            # --- COOLER 5 --- 

            # for constants and flows
            print("\n\033[94mNow entering for Cooler 5...\033[0m\n")
            if status[4].strip().upper() == 'OOS' or status[4].strip().upper() == 'STANDBY':
                print("\033[91mError: Comp 5 OOS...\nUnable to compute duties. Exiting program...\033[0m")  
                sys.exit()
            else:
                volHGF = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQK7wDAATkRQ" \
                    "SURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpIQVdLRVlFLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MR" \
                    "VRSRUNJUDUuQ0FMQ1VMQVRFRkxPV1lEWV80QU0uRlFJVDc1MjUwMQ/value",
                    auth=(username, password), verify=False).json() # gets pi data (flowrate and timestamp) for Cooler 5
                if not isinstance(volHGF["Value"], (int, float)): # if value has bad input
                    print("\033[91mError: Bad transmitter input for Cooler 5 flow...\nDuty marked as 0 BTU/hr...\033[0m")  
                    mHGF = 0
                else:
                    mHGF = volHGF["Value"] * 1/397.3 * mmHGF / 24 # gets mass flow rate [1000-lb/hr]

                # -- INTERCOOLER 1 --
                # gaining values for q computation
                T_in = (T_in_list[T_in_list['Label'] == '1ST STAGE #1 NORTH OUTLET TEMP'].iloc[4])['Response']
                T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQgqgBAATkRQSU" \
                    "RBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpIQVdLRVlFLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRS" \
                    "RUNJUDUuRklSU1RTVEFHRUlOVEVSQ09PTEVSVEVNUC5USVQ3NTI1MTE/value",
                    auth=(username, password), verify=False).json() # gets pi data for outlet temp for Cooler 5, IC-1
                if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                    print("\033[91mError: Bad transmitter input for Cooler 5 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
                    q_C5_IC1 = 0
                else:
                    q_C5_IC1 = mHGF * cpHGF * (T_in - T_out["Value"]) # heat rejection (1000-BTU/hr)

                # -- INTERCOOLER 2 --

                # gaining values for q computation
                T_in = (T_in_list[T_in_list['Label'] == '2ND STAGE OUTLET TEMP'].iloc[1])['Response']
                T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQg6gBAATkRQ" \
                    "SURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpIQVdLRVlFLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5" \
                    "MRVRSRUNJUDUuU0VDT05EU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUNzUyNTEy/value",
                    auth=(username, password), verify=False).json() # gets pi data for outlet temp for Cooler 5, IC-2
                if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                    print("\033[91mError: Bad transmitter input for Cooler 5 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
                    q_C5_IC2 = 0
                else:
                    q_C5_IC2 = mHGF * cpHGF * (T_in - T_out["Value"]) # heat rejection (1000-BTU/hr)

                # -- INTERCOOLER 3 --

                # gaining values for q computation
                T_in = (T_in_list[T_in_list['Label'] == '3RD STAGE OUTLET TEMP'].iloc[4])['Response']
                T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQhKgBAATkRQS" \
                    "URBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpIQVdLRVlFLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MR" \
                    "VRSRUNJUDUuVEhJUkRTVEFHRUlOVEVSQ09PTEVSVEVNUC5USVQ3NTI1MTM/value",
                    auth=(username, password), verify=False).json() # gets pi data for outlet temp for Cooler 5, IC-3
                if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                    print("\033[91mError: Bad transmitter input for Cooler 5 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
                    q_C5_IC3 = 0
                else:
                    q_C5_IC3 = mHGF * cpHGF * (T_in - T_out["Value"]) # heat rejection (1000-BTU/hr)

                # -- INTERCOOLER 4 --

                # gaining values for q computation
                T_in = (T_in_list[T_in_list['Label'] == '4TH STAGE OUTLET TEMP'].iloc[4])['Response']
                T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQhagBAATkRQ" \
                    "SURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpIQVdLRVlFLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MR" \
                    "VRSRUNJUDUuRk9VUlRIU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUNzUyNTE0/value",
                    auth=(username, password), verify=False).json() # gets pi data for outlet temp for Cooler 5, IC-4
                if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                    print("\033[91mError: Bad transmitter input for Cooler 5 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
                    q_BB2C5_IC4 = 0
                else:
                    q_C5_IC4 = mHGF * cpHGF * (T_in - T_out["Value"]) # heat rejection (1000-BTU/hr)
            
                # EJW - is using flows from data sheet

                T_in = (T_in_list[T_in_list['Label'] == 'ENGINE: JACKET WATER OUTLET TEMP'].iloc[4])['Response']
                T_out = (T_in_list[T_in_list['Label'] == 'ENGINE: JACKET WATER INLET TEMP'].iloc[4])['Response']
                q_C5_EJW = ((575 * 192.5 * 1/24) * (1.02 * 60.793))  * (0.643) * (T_in - T_out) / 1000 # heat rejection (1000-BTU/hr)

            # -- END OF INLET GAS COOLERS --
            
            break
        elif cooler == 'flash gas' or cooler == '1': # if hgf fg coolers selected
            print("You have selected HGF Flash Gas Coolers")

            print("\033[91mNOTE: FG3, and TAW for all units unable to be computed at this time\033[0m\n")
            
            # ---- FLASH GAS ----
            input("Verify that the operator rounds data is uploaded to this folder with the name 'FGRounds.xlsx' and press ENTER to continue...\n")
            T_in_list = pd.read_excel("FGRounds.xlsx")[['Label', 'Response']]
            status = (T_in_list[T_in_list['Label'] == 'COMPRESSOR STATUS'])['Response'].tolist()
            if status[0].strip().upper() == 'OOS' or status[0].strip().upper() == 'STANDBY':
                print("\033[91mError: FG Comp 1 OOS...\nUnable to compute duties. Exiting program...\033[0m")  
                sys.exit()
            else:
                T_in_list['Response'] = pd.to_numeric((T_in_list['Response'].astype(str).str.replace(' degF', '', regex=False)), errors='coerce')

                # --- COOLER 1 --- 
                print("\033[94mNow entering for Cooler 1...\033[0m\n")

                # -- INTERCOOLER --
                # gaining values for q computation
                volHGF = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQ46gBAATkR" \
                    "QSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpIQVdLRVlFLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uRkx" \
                    "BU0hHQVMxLkZJUlNUU1RBR0VJTkxFVEZMT1dSQVRFLkZJVDc1ODExMQ/value",
                    auth=(username, password), verify=False).json() # gets pi data (flowrate and timestamp) for Cooler 1
                if not isinstance(volHGF["Value"], (int, float)): # if value has bad input
                    print("\033[91mError: Bad transmitter input for Cooler FG1 flow...\nDuty marked as 0 BTU/hr...\033[0m")  
                    mHGF = 0
                else:
                    mHGF = volHGF["Value"] * 1/397.3 * mmHGF / 24 # gets mass flow rate [1000-lb/hr]

                T_in = (T_in_list[T_in_list['Label'] == 'CYLINDER 1 OUTLET'].iloc[0])['Response']
                T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQt6cBAATkRQ" \
                    "SURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpIQVdLRVlFLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uRkxB" \
                    "U0hHQVMxLkVOR0lORUVYSEFVU1RURU1QLlRJVDc1ODExMw/value",
                    auth=(username, password), verify=False).json() # gets pi data for outlet temp for Cooler 1, IC-1
                if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                    print("\033[91mError: Bad transmitter input for FG1 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
                    q_FG1_IC1 = 0
                else:
                    q_FG1_IC1 = mHGF * cpHGF * (T_in - T_out["Value"] / 100) # heat rejection (1000-BTU/hr) -- NOTE: T_out DIVIDED BY 100 BC PI TAG VALUE INACCURATE

                # EJW
                
                T_in = float(input("Input Cooler \033[91mInlet Temp\033[0m for EJW (°F): \n")) # temp into cooler (hot) - NOT AUTOMATED
                T_out = float(input("Input Cooler \033[94mOutlet Temp\033[0m for EJW (°F): \n")) # temp into cooler (hot) - NOT AUTOMATED
                q_FG1_EJW = ((575 * 192.5 * 1/24) * (1.02 * 60.793))  * (0.643) * (T_in - T_out) / 1000 # heat rejection (1000-BTU/hr)

                '''
                NOTE: TAW COMMENTED OUT - INSUFFICIENT INFORMATION/WORTHWHILE INSIGHT
                # -- TAW --
                volTAW = values from data sheet
                mTAW = other values

                # gaining values for q computation
                T_in = float(input("Input Cooler Inlet Temp for TAW (°F): \n")) # temp into cooler (hot) - NOT AUTOMATED
                T_out = float(input("Input Cooler Outlet Temp for TAW (°F): \n")) # temp out of cooler (hot) - NOT AUTOMATED
                q_FG1_TAW = mTAW * cpTAW * (T_in - T_out) # heat rejection (1000-BTU/hr)
                '''

                '''
            # --- COOLER 2 --- 
            print("\033[94mNow entering for Cooler 2...\033[0m\n")
            if status[1].strip().upper() == 'OOS' or status[1].strip().upper() == 'STANDBY':
                print("\033[91mError: FG Comp 2 OOS...\nUnable to compute duties. Exiting program...\033[0m")  
                sys.exit()
            else:
                # -- INTERCOOLER --
                # gaining values for q computation
                volHGF = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQ6KgBAATkR" \
                    "QSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpIQVdLRVlFLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uRkx" \
                    "BU0hHQVMyLkZJUlNUU1RBR0VJTkxFVEZMT1dSQVRFLkZJVDc1ODIxMQ/value",
                    auth=(username, password), verify=False).json() # gets pi data (flowrate and timestamp) for Cooler 2
                if not isinstance(volHGF["Value"], (int, float)): # if value has bad input
                    print("\033[91mError: Bad transmitter input for FG2 flow...\nDuty marked as 0 BTU/hr...\033[0m")  
                    mHGF = 0
                else:
                    mHGF = volHGF["Value"] * 1/397.3 * mmHGF / 24 # gets mass flow rate [1000-lb/hr]

                T_in = float(input("Input Cooler \033[91mInlet Temp\033[0m for IC-1 (°F): \n")) # temp into cooler (hot) - NOT AUTOMATED
                T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQw6cBAATkRQS" \
                    "URBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpIQVdLRVlFLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uRkxBU0h" \
                    "HQVMyLkVOR0lORUVYSEFVU1RURU1QLlRJVDc1ODIxMw/value",
                    auth=(username, password), verify=False).json() # gets pi data for outlet temp for Cooler 1, IC-1
                if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                    print("\033[91mError: Bad transmitter input for Cooler FG2 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
                    q_FG2_IC1 = 0
                else:
                    q_FG2_IC1 = mHGF * cpHGF * (T_in - T_out["Value"]) # heat rejection (1000-BTU/hr)

                # EJW - is using flows from datasheet

                T_in = float(input("Input Cooler \033[91mInlet Temp\033[0m for EJW (°F): \n")) # temp into cooler (hot) - NOT AUTOMATED
                T_out = float(input("Input Cooler \033[94mOutlet Temp\033[0m for EJW (°F): \n")) # temp into cooler (hot) - NOT AUTOMATED
                q_FG2_EJW = ((575 * 192.5 * 1/24) * (1.02 * 60.793))  * (0.643) * (T_in - T_out) / 1000 # heat rejection (1000-BTU/hr)
                '''
                
                '''
                NOTE: TAW COMMENTED OUT - INSUFFICIENT INFORMATION/WORTHWHILE INSIGHT
                # -- TAW --

                # gaining values for q computation
                T_in = float(input("Input Cooler Inlet Temp for TAW (°F): \n")) # temp into cooler (hot) - NOT AUTOMATED
                T_out = float(input("Input Cooler Outlet Temp for TAW (°F): \n")) # temp out of cooler (hot) - NOT AUTOMATED
                q_FG2_TAW = mTAW * cpTAW * (T_in - T_out) # heat rejection (1000-BTU/hr)

                '''

            '''
            # --- COOLER 3 --- 
            print("\033[94mNow entering for Cooler 3...\033[0m\n")

            # -- INTERCOOLER --
            # gaining values for q computation
            volHGF = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQ6qgBAATkRQ" \
                "SURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpIQVdLRVlFLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uRkxBU0" \
                "hHQVMzLkZJUlNUU1RBR0VJTkxFVEZMT1dSQVRFLkZJVDc1ODMxMQ/value",
                auth=(username, password), verify=False).json() # gets pi data (flowrate and timestamp) for Cooler 3
            if not isinstance(volHGF["Value"], (int, float)): # if value has bad input
                print("\033[91mError: Bad transmitter input for FG2 flow...\nDuty marked as 0 BTU/hr...\033[0m")  
                mHGF = 0
            else:
                mHGF = volHGF["Value"] * 1/397.3 * mmHGF / 24 # gets mass flow rate [1000-lb/hr] 

            T_in = float(input("Input Cooler Inlet Temp for IC-1 (°F): \n")) # temp into cooler (hot) - NOT AUTOMATED
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQz6cBAATkRQ" \
                "SURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpIQVdLRVlFLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uRkxB" \
                "U0hHQVMzLkVOR0lORUVYSEFVU1RURU1QLlRJVDc1ODMxMw/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp for Cooler 3, IC-1
            if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                print("\033[91mError: Bad transmitter input for FG3 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_FG3_IC1 = 0
            else:
                q_FG3_IC1 = mHGF * cpHGF * (T_in - T_out["Value"]) # heat rejection (1000-BTU/hr)
            
            # EJW - is using flows from data sheet

            T_in = float(input("Input Cooler \033[91mInlet Temp\033[0m for EJW (°F): \n")) # temp into cooler (hot) - NOT AUTOMATED
            T_out = float(input("Input Cooler \033[94mOutlet Temp\033[0m for EJW (°F): \n")) # temp into cooler (hot) - NOT AUTOMATED
            q_FG3_EJW = ((575 * 192.5 * 1/24) * (1.02 * 60.793))  * (0.643) * (T_in - T_out) / 1000 # heat rejection (1000-BTU/hr)
            '''

            '''
            NOTE: TAW COMMENTED OUT - INSUFFICIENT INFORMATION/WORTHWHILE INSIGHT
            # -- TAW --

            # gaining values for q computation
            T_in = float(input("Input Cooler Inlet Temp for TAW (°F): \n")) # temp into cooler (hot) - NOT AUTOMATED
            T_out = float(input("Input Cooler Outlet Temp for TAW (°F): \n")) # temp out of cooler (hot) - NOT AUTOMATED
            q_FG3_TAW = mTAW * cpTAW * (T_in - T_out) # heat rejection (1000-BTU/hr)

            '''

            # -- END OF FLASH GAS COOLERS --

            break
        elif cooler == 'refrig' or cooler == '2':
            print("You have selected HGF Refrig Coolers\n")

            input("Verify that the operator rounds data is uploaded to this folder with the name 'RefrigRounds.xlsx' and press ENTER to continue...\n")
            T_in_list = pd.read_excel("RefrigRounds.xlsx")[['Label', 'Response']]
            
            # --- COOLER 1 ---
            print("\033[94mNow entering for Cooler 1...\033[0m\n")
            status = (T_in_list[T_in_list['Label'] == 'COMPRESSOR STATUS'])['Response'].tolist()
            if status[0].strip().upper() == 'OOS' or status[0].strip().upper() == 'STANDBY':
                print("\033[91mError: FG Comp 1 OOS...\nUnable to compute duties. Exiting program...\033[0m")  
                sys.exit()
            else:
                T_in_list['Response'] = pd.to_numeric((T_in_list['Response'].astype(str).str.replace(' degF', '', regex=False)), errors='coerce')

                # -- EJW --

                # gaining values for q computation
                # EJW

                T_in = (T_in_list[T_in_list['Label'] == 'ENGINE: JACKET WATER OUTLET TEMP'].iloc[0])['Response']
                T_out = (T_in_list[T_in_list['Label'] == 'ENGINE: JACKET WATER INLET TEMP'].iloc[0])['Response']
                q_Re1_EJW = ((575 * 192.5 * 1/24) * (1.02 * 60.793))  * (0.643) * (T_in - T_out) / 1000 # heat rejection (1000-BTU/hr)

                '''
                NOTE: TAW COMMENTED OUT - INSUFFICIENT INFORMATION/WORTHWHILE INSIGHT
                # -- TAW --

                # gaining values for q computation
                T_in = float(input("Input Cooler Inlet Temp for TAW (°F): \n")) # temp into cooler (hot) - NOT AUTOMATED
                T_out = float(input("Input Cooler Inlet Temp for TAW (°F): \n")) # temp into cooler (hot) - NOT AUTOMATED
                q_Re1_TAW = [insert taw mass] * [insert taw cp] * (T_in - T_out) # heat rejection (1000-BTU/hr)
                '''

                '''
            # --- COOLER 2 --- 
            print("\n\033[94mNow entering for Cooler 2...\033[0m\n")

            # -- EJW --

            # gaining values for q computation

            # EJW

            T_in = float(input("Input Cooler \033[91mInlet Temp\033[0m for EJW (°F): \n")) # temp into cooler (hot) - NOT AUTOMATED
            T_out = float(input("Input Cooler \033[94mOutlet Temp\033[0m for EJW (°F): \n")) # temp into cooler (hot) - NOT AUTOMATED
            q_Re2_EJW = ((575 * 192.5 * 1/24) * (1.02 * 60.793))  * (0.643) * (T_in - T_out) / 1000 # heat rejection (1000-BTU/hr)
            '''

            '''
            # -- TAW --

            # gaining values for q computation
            T_in = float(input("Input Cooler Inlet Temp for TAW (°F): \n")) # temp into cooler (hot) - NOT AUTOMATED
            T_out = float(input("Input Cooler Inlet Temp for TAW (°F): \n")) # temp into cooler (hot) - NOT AUTOMATED
            q_Re2_TAW = [insert taw mass] * [insert taw cp] * (T_in - T_out) # heat rejection (1000-BTU/hr)
            '''

            # -- END OF refrig COOLERS --

            break
        else:
            print("Unknown cooler type. Please enter 'Inlet Gas', 'Flash Gas', or 'Refrig'")

# BBCS2 coolers code    
elif station == "BBCS2" or station == '1':
    print("You have selected BBCS2\n")
    cpBB2 = 0.5436823 # Cp of inlet gas [BTU/lb/°F]
    mmBB2 = 25.44 # molecular mass of inlet gas [lb/lb-mol]

    while True:        
        print("\033[94mNow generating cooler analysis of Coolers 1-10...\033[0m\n")
        
        # --- COOLER 1 --- 
        volBB2 = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQRMICAATkRQSUR" \
            "BQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOT" \
            "EVUUkVDSVAxLkNBTENVTEFURURGTE9XLktNNzkyMTIw/value",
            auth=(username, password), verify=False).json() # gets pi data (flowrate and timestamp) for Cooler 1
        if not isinstance(volBB2["Value"], (int, float)): # if value has bad input
                print("\033[91mError: Bad transmitter input for Cooler 1 flow...\nDuty marked as 0 BTU/hr...\033[0m")  
                mBB2 = 0
        else:
            mBB2 = volBB2["Value"] * 1/397.3 * mmBB2 / 24 * 1000 # gets mass flow rate [1000-lb/hr]

        # -- INTERCOOLER 1 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSk" \
            "CZuCntY8tYTQTcICAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTF" \
            "VFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVAxLkZJUlNUU1RBR0VJTkxFVFRFTVAuVElUNzkyMTA4QQ/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
            print("\033[91mError: Bad transmitter input for Cooler 1 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BB2C1_IC1 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQUsICAATkRQSURBQ09MTFxVU0EuTkQ" \
                "uTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVAxLkZJUlNUU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUNzkyMTEy/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 1 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C1_IC1 = 0
            else:
                q_BB2C1_IC1 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 2 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZ" \
            "uCntY8tYTQT8ICAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQl" \
            "VUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVAxLlNFQ09ORFNUQUdFT1VUTEVUVEVNUC5USVQ3OTIxMDk/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
            print("\033[91mError: Bad transmitter input for Cooler 1 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BB2C1_IC2 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQU8ICAATkRQSURBQ09MTFxV" \
                "U0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVAxLlNFQ09ORFNUQUdFSU5URVJDT09MRVJURU1QLlRJVDc5MjExMw/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 1 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C1_IC2 = 0
            else:
                q_BB2C1_IC2 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 3 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZ" \
            "uCntY8tYTQUMICAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQl" \
            "VUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVAxLlRISVJEU1RBR0VPVVRMRVRURU1QLlRJVDc5MjExMA/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
            print("\033[91mError: Bad transmitter input for Cooler 1 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BB2C1_IC3 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQVMICAATkRQSURBQ0" \
                "9MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVAxLlRISVJEU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUNzkyMTE0/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 1 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C1_IC3 = 0
            else:
                q_BB2C1_IC3 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 4 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCn" \
            "tY8tYTQUcICAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEV" \
            "TMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVAxLkZPVVJUSFNUQUdFT1VUTEVUVEVNUC5USVQ3OTIxMTE/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
            print("\033[91mError: Bad transmitter input for Cooler 1 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BB2C1_IC4 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQVcICAATkRQSU" \
                "RBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVAxLkZPVVJUSFNUQUdFSU5URVJDT09MRVJURU1QLlRJVDc5MjExNQ/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 1 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C1_IC4 = 0
            else:
                q_BB2C1_IC4 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # --- COOLER 2 --- 
        volBB2 = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQccICAATkR" \
            "QSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OL" \
            "klOTEVUUkVDSVAyLkNBTENVTEFURURGTE9XLktNNzkyMjIw/value",
            auth=(username, password), verify=False).json() # gets pi data (flowrate and timestamp) for Cooler 2
        if not isinstance(volBB2["Value"], (int, float)):
            print("\033[91mError: Bad transmitter input for Cooler 2 flow...\nDuty marked as 0 BTU/hr...\033[0m")  
            mBB2 = 0
        else:
            mBB2 = volBB2["Value"] * 1/397.3 * mmBB2 / 24 * 1000 # gets mass flow rate [1000-lb/hr]

        # -- INTERCOOLER 1 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8t" \
            "YTQgcICAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk" \
            "9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVAyLkZJUlNUU1RBR0VJTkxFVFRFTVAuVElUNzkyMjA4QQ/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
            print("\033[91mError: Bad transmitter input for Cooler 2 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BB2C2_IC1 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQg8ICAAT" \
                "kRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVAyLkZJUlNUU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUNzkyMjEy/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 2 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C2_IC1 = 0
            else:
                q_BB2C2_IC1 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 2 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSk" \
            "CZuCntY8tYTQr8ICAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVF" \
            "QlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVAyLlNFQ09ORFNUQUdFT1VUTEVUVEVNUC5USVQ3OTIyMDk/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
            print("\033[91mError: Bad transmitter input for Cooler 2 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BB2C2_IC2 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQhMICAATkRQSURBQ09" \
                "MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVAyLlNFQ09ORFNUQUdFSU5URVJDT09MRVJURU1QLlRJVDc5MjIxMw/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 2 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C2_IC2 = 0
            else:
                q_BB2C2_IC2 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 3 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQtsICAA" \
            "TkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPT" \
            "VBSRVNTSU9OLklOTEVUUkVDSVAyLlRISVJEU1RBR0VPVVRMRVRURU1QLlRJVDc5MjIxMA/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
            print("\033[91mError: Bad transmitter input for Cooler 2 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BB2C2_IC3 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQhcICAATkRQS" \
                "URBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVAyLlRISVJEU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUNzkyMjE0/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 2 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C2_IC3 = 0
            else:
                q_BB2C2_IC3 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 4 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQsMICAATkRQSURBQ09MTFxVU0E" \
            "uTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVAyLkZPVVJUSFNUQUdFT1VUTEVUVEVNUC5USVQ3OTIyMTE/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
            print("\033[91mError: Bad transmitter input for Cooler 2 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BB2C2_IC4 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQhsICAATkRQSU" \
                "RBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVAyLkZPVVJUSFNUQUdFSU5URVJDT09MRVJURU1QLlRJVDc5MjIxNQ/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 2 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C2_IC4 = 0
            else:
                q_BB2C2_IC4 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # --- COOLER 3 --- 
        volBB2 = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQnsICAATk" \
            "RQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVN" \
            "TSU9OLklOTEVUUkVDSVAzLkNBTENVTEFURURGTE9XLktNNzkyMzIw/value",
            auth=(username, password), verify=False).json() # gets pi data (flowrate and timestamp) for Cooler 3
        if not isinstance(volBB2["Value"], (int, float)):
            print("\033[91mError: Bad transmitter input for Cooler 3 flow...\nDuty marked as 0 BTU/hr...\033[0m")  
            mBB2 = 0
        else:
            mBB2 = volBB2["Value"] * 1/397.3 * mmBB2 / 24 * 1000 # gets mass flow rate [1000-lb/hr]

        # -- INTERCOOLER 1 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQsc" \
            "ICAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLk" \
            "dBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVAzLkZJUlNUU1RBR0VJTkxFVFRFTVAuVElUNzkyMzA4QQ/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
            print("\033[91mError: Bad transmitter input for Cooler 3 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BB2C3_IC1 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQt8ICAATkR" \
                "QSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVAzLkZJUlNUU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUNzkyMzEy/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 3 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C3_IC1 = 0
            else:
                q_BB2C3_IC1 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 2 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQs8ICA" \
            "ATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0" \
            "NPTVBSRVNTSU9OLklOTEVUUkVDSVAzLlNFQ09ORFNUQUdFT1VUTEVUVEVNUC5USVQ3OTIzMDk/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
            print("\033[91mError: Bad transmitter input for Cooler 3 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BB2C3_IC2 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQuMICAA" \
                "TkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVAzLlNFQ09ORFNUQUdFSU5URVJDT09MRVJURU1QLlRJVDc5MjMxMw/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 3 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C3_IC2 = 0
            else:
                q_BB2C3_IC2 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 3 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQtMICAAT" \
            "kRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0N" \
            "PTVBSRVNTSU9OLklOTEVUUkVDSVAzLlRISVJEU1RBR0VPVVRMRVRURU1QLlRJVDc5MjMxMA/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
            print("\033[91mError: Bad transmitter input for Cooler 3 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BB2C3_IC3 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQucICAATkRQ" \
                "SURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVAzLlRISVJEU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUNzkyMzE0/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 3 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C3_IC1 = 0
            else:
                q_BB2C3_IC3 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 4 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQtcICAAT" \
            "kRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0" \
            "NPTVBSRVNTSU9OLklOTEVUUkVDSVAzLkZPVVJUSFNUQUdFT1VUTEVUVEVNUC5USVQ3OTIzMTE/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
            print("\033[91mError: Bad transmitter input for Cooler 3 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BB2C3_IC4 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQusICAATk" \
                "RQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVAzLkZPVVJUSFNUQUdFSU5URVJDT09MRVJURU1QLlRJVDc5MjMxNQ/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 3 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C3_IC4 = 0
            else:
                q_BB2C3_IC4 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # --- COOLER 4 --- 
        volBB2 = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQ0cICAATkRQ" \
            "SURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLk" \
            "lOTEVUUkVDSVA0LkNBTENVTEFURURGTE9XLktNNzkyNDIw/value",
            auth=(username, password), verify=False).json() # gets pi data (flowrate and timestamp) for Cooler 4
        if not isinstance(volBB2["Value"], (int, float)):
            print("\033[91mError: Bad transmitter input for Cooler 4 flow...\nDuty marked as 0 BTU/hr...\033[0m")  
            mBB2 = 0
        else:
            mBB2 = volBB2["Value"] * 1/397.3 * mmBB2 / 24 * 1000 # gets mass flow rate [1000-lb/hr]

        # -- INTERCOOLER 1 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQ4cICA" \
            "ATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0" \
            "NPTVBSRVNTSU9OLklOTEVUUkVDSVA0LkZJUlNUU1RBR0VJTkxFVFRFTVAuVElUNzkyNDA4QQ/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
            print("\033[91mError: Bad transmitter input for Cooler 4 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BB2C4_IC1 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQ5sICAATkRQS" \
                "URBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVA0LkZJUlNUU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUNzkyNDEy/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 4 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C3_IC1 = 0
            else:
                q_BB2C4_IC1 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 2 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQ48ICAAT" \
            "kRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVB" \
            "SRVNTSU9OLklOTEVUUkVDSVA0LlNFQ09ORFNUQUdFT1VUTEVUVEVNUC5USVQ3OTI0MDk/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
            print("\033[91mError: Bad transmitter input for Cooler 4 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BB2C4_IC2 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQ58ICAATkRQSURB" \
                "Q09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVA0LlNFQ09ORFNUQUdFSU5URVJDT09MRVJURU1QLlRJVDc5MjQxMw/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 4 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C4_IC2 = 0
            else:
                q_BB2C4_IC2 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 3 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQ5MICAAT" \
            "kRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBS" \
            "RVNTSU9OLklOTEVUUkVDSVA0LlRISVJEU1RBR0VPVVRMRVRURU1QLlRJVDc5MjQxMA/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
            print("\033[91mError: Bad transmitter input for Cooler 4 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BB2C4_IC3 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQ6MICAATkRQSU" \
                "RBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVA0LlRISVJEU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUNzkyNDE0/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 4 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C4_IC3 = 0
            else:
                q_BB2C4_IC3 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 4 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQ5cICA" \
            "ATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0" \
            "NPTVBSRVNTSU9OLklOTEVUUkVDSVA0LkZPVVJUSFNUQUdFT1VUTEVUVEVNUC5USVQ3OTI0MTE/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
            print("\033[91mError: Bad transmitter input for Cooler 4 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BB2C4_IC4 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQ6cICAATkRQ" \
                "SURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVA0LkZPVVJUSFNUQUdFSU5URVJDT09MRVJURU1QLlRJVDc5MjQxNQ/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 4 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C4_IC4 = 0
            else:
                q_BB2C4_IC4 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # --- COOLER 5 --- 
        volBB2 = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQ_sICAATkRQ" \
            "SURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9O" \
            "LklOTEVUUkVDSVA1LkNBTENVTEFURURGTE9XLktNNzkyNTIw/value",
            auth=(username, password), verify=False).json() # gets pi data (flowrate and timestamp) for Cooler 5
        if not isinstance(volBB2["Value"], (int, float)):
            print("\033[91mError: Bad transmitter input for Cooler 5 flows...\nDuty marked as 0 BTU/hr...\033[0m")  
            mBB2 = 0
        else:
            mBB2 = volBB2["Value"] * 1/397.3 * mmBB2 / 24 * 1000 # gets mass flow rate [1000-lb/hr]

        # -- INTERCOOLER 1 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQDsMCAATkRQSURBQ" \
            "09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOT" \
            "EVUUkVDSVA1LkZJUlNUU1RBR0VJTkxFVFRFTVAuVElUNzkyNTA4QQ/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
            print("\033[91mError: Bad transmitter input for Cooler 5 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BB2C5_IC1 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQE8MCAATkRQSURB" \
                "Q09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVA1LkZJUlNUU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUNzkyNTEy/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 5 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C5_IC1 = 0
            else:
                q_BB2C5_IC1 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 2 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQEMMCAATk" \
            "RQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTV" \
            "BSRVNTSU9OLklOTEVUUkVDSVA1LlNFQ09ORFNUQUdFT1VUTEVUVEVNUC5USVQ3OTI1MDk/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
            print("\033[91mError: Bad transmitter input for Cooler 5 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BB2C5_IC2 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQFMMCAATkR" \
                "QSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVA1LlNFQ09ORFNUQUdFSU5URVJDT09MRVJURU1QLlRJVDc5MjUxMw/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 5 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C5_IC2 = 0
            else:
                q_BB2C5_IC2 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 3 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQEcMCAA" \
            "TkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVB" \
            "SRVNTSU9OLklOTEVUUkVDSVA1LlRISVJEU1RBR0VPVVRMRVRURU1QLlRJVDc5MjUxMA/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
            print("\033[91mError: Bad transmitter input for Cooler 5 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BB2C5_IC3 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQFcMCAATkRQSU" \
                "RBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVA1LlRISVJEU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUNzkyNTE0/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 5 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C5_IC3 = 0
            else:
                q_BB2C5_IC3 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 4 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQEsMCAATkRQS" \
            "URBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTS" \
            "U9OLklOTEVUUkVDSVA1LkZPVVJUSFNUQUdFT1VUTEVUVEVNUC5USVQ3OTI1MTE/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
            print("\033[91mError: Bad transmitter input for Cooler 5 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BB2C5_IC4 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQFsMCAATkRQSURBQ0" \
                "9MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVA1LkZPVVJUSFNUQUdFSU5URVJDT09MRVJURU1QLlRJVDc5MjUxNQ/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 5 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C5_IC4 = 0
            else:
                q_BB2C5_IC4 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # --- COOLER 6 --- 
        volBB2 = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQLcMCAAT" \
            "kRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNT" \
            "SU9OLklOTEVUUkVDSVA2LkNBTENVTEFURURGTE9XLktNNzkyNjIw/value",
            auth=(username, password), verify=False).json() # gets pi data (flowrate and timestamp) for Cooler 6
        if not isinstance(volBB2["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 6 flow...\nDuty marked as 0 BTU/hr...\033[0m")  
                mBB2 = 0
        else:
            mBB2 = volBB2["Value"] * 1/397.3 * mmBB2 / 24 * 1000 # gets mass flow rate [1000-lb/hr]

        # -- INTERCOOLER 1 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQPcMCAATkRQSUR" \
            "BQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU" \
            "9OLklOTEVUUkVDSVA2LkZJUlNUU1RBR0VJTkxFVFRFTVAuVElUNzkyNjA4QQ/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 6 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C6_IC1 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQQsMCAATkRQSURBQ09MTFxV" \
                "U0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVA2LkZJUlNUU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUNzkyNjEy/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 6 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C6_IC1 = 0
            else:
                q_BB2C6_IC1 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 2 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQP8MCAATkRQ" \
            "SURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSR" \
            "VNTSU9OLklOTEVUUkVDSVA2LlNFQ09ORFNUQUdFT1VUTEVUVEVNUC5USVQ3OTI2MDk/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 6 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C6_IC2 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQQ8MCAATkRQSURB" \
                "Q09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVA2LlNFQ09ORFNUQUdFSU5URVJDT09MRVJURU1QLlRJVDc5MjYxMw/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 6 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C6_IC2 = 0
            else:
                q_BB2C6_IC2 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 3 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQQMMCAATkRQSU" \
            "RBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OL" \
            "klOTEVUUkVDSVA2LlRISVJEU1RBR0VPVVRMRVRURU1QLlRJVDc5MjYxMA/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 6 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C6_IC3 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQRMMCAATkRQSURB" \
                "Q09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVA2LlRISVJEU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUNzkyNjE0/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 6 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C6_IC3 = 0
            else:
                q_BB2C6_IC3 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 4 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQQcMCAATkRQSURBQ09MT" \
            "FxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTE" \
            "VUUkVDSVA2LkZPVVJUSFNUQUdFT1VUTEVUVEVNUC5USVQ3OTI2MTE/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 6 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C6_IC4 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQRcMCAATkRQSURBQ09M" \
                "TFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVA2LkZPVVJUSFNUQUdFSU5URVJDT09MRVJURU1QLlRJVDc5MjYxNQ/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 6 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C6_IC4 = 0
            else:
                q_BB2C6_IC4 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # --- COOLER 7 --- 
        volBB2 = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQW8MCAATkRQ" \
            "SURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVA3LkNBTENVTEFURURGTE9XLktNNzkyNzIw/value",
            auth=(username, password), verify=False).json() # gets pi data (flowrate and timestamp) for Cooler 7
        if not isinstance(volBB2["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 7 flow...\nDuty marked as 0 BTU/hr...\033[0m")  
                mBB2 = 0
        else:
            mBB2 = volBB2["Value"] * 1/397.3 * mmBB2 / 24 * 1000 # gets mass flow rate [1000-lb/hr]

        # -- INTERCOOLER 1 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQbMMCAATkRQSURBQ09MT" \
            "FxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkV" \
            "DSVA3LkZJUlNUU1RBR0VJTkxFVFRFTVAuVElUNzkyNzA4QQ/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 7 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C7_IC1 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQccMCAATkRQSURBQ09MTF" \
                "xVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVA3LkZJUlNUU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUNzkyNzEy/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 7 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C7_IC1 = 0
            else:
                q_BB2C7_IC1 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 2 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQbsMCAATkRQSURBQ09M" \
            "TFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEV" \
            "UUkVDSVA3LlNFQ09ORFNUQUdFT1VUTEVUVEVNUC5USVQ3OTI3MDk/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 7 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C7_IC2 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQcsMCAATkRQSURBQ09MTF" \
                "xVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVA3LlNFQ09ORFNUQUdFSU5URVJDT09MRVJURU1QLlRJVDc5MjcxMw/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                    print("\033[91mError: Bad transmitter input for Cooler 7 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
                    q_BB2C7_IC2 = 0
            else:
                q_BB2C7_IC2 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 3 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQb8MCAATkRQSURBQ09MTFxVU0Eu" \
            "TkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUk" \
            "VDSVA3LlRISVJEU1RBR0VPVVRMRVRURU1QLlRJVDc5MjcxMA/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 7 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C7_IC3 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQc8MCAATkRQSURBQ09MTFxVU0E" \
                "uTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVA3LlRISVJEU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUNzkyNzE0/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                    print("\033[91mError: Bad transmitter input for Cooler 7 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
                    q_BB2C6_IC1 = 0
            else:
                q_BB2C7_IC3 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 4 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQcMMCAATkRQSURBQ09MTFxVU0EuTkQuTU" \
            "lELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVA3LkZPVVJUSFNUQUdFT1V" \
            "UTEVUVEVNUC5USVQ3OTI3MTE/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 7 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C7_IC4 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQdMMCAATkRQSURBQ09MTFxVU0EuTkQuTU" \
                "lELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVA3LkZPVVJUSFNUQUdFSU5URVJDT09MRVJURU1QLlRJVDc5MjcxNQ/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 7 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C7_IC4 = 0
            else:
                q_BB2C7_IC4 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # --- COOLER 8 --- 
        volBB2 = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQhMMCAATkRQ" \
            "SURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTS" \
            "U9OLklOTEVUUkVDSVA4LkNBTENVTEFURURGTE9XLktNNzkyODIw/value",
            auth=(username, password), verify=False).json() # gets pi data (flowrate and timestamp) for Cooler 8
        if not isinstance(volBB2["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 8 flow...\nDuty marked as 0 BTU/hr...\033[0m")  
                mBB2 = 0
        else:
            mBB2 = volBB2["Value"] * 1/397.3 * mmBB2 / 24 * 1000 # gets mass flow rate [1000-lb/hr]

        # -- INTERCOOLER 1 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQlcMCAATkRQSURBQ" \
            "09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OL" \
            "klOTEVUUkVDSVA4LkZJUlNUU1RBR0VJTkxFVFRFTVAuVElUNzkyODA4QQ/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 8 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C8_IC1 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQmsMCAATkRQSURBQ09M" \
                "TFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVA4LkZJUlNUU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUNzkyODEy/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 8 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C8_IC1 = 0
            else:
                q_BB2C8_IC1 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 2 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQl8MCAATkRQ" \
            "SURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRV" \
            "NTSU9OLklOTEVUUkVDSVA4LlNFQ09ORFNUQUdFT1VUTEVUVEVNUC5USVQ3OTI4MDk/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 8 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C8_IC2 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQm8MCAATkRQSURBQ0" \
                "9MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVA4LlNFQ09ORFNUQUdFSU5URVJDT09MRVJURU1QLlRJVDc5MjgxMw/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 8 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C8_IC2 = 0
            else:
                q_BB2C8_IC2 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 3 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQmMMCAATkRQSURB" \
            "Q09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9O" \
            "LklOTEVUUkVDSVA4LlRISVJEU1RBR0VPVVRMRVRURU1QLlRJVDc5MjgxMA/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 8 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C8_IC3 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQnMMCAATkRQSURBQ09" \
                "MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVA4LlRISVJEU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUNzkyODE0/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_in["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 8 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C8_IC3 = 0
            else:
                q_BB2C8_IC3 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 4 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQmcMCAATkRQSUR" \
            "BQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNT" \
            "SU9OLklOTEVUUkVDSVA4LkZPVVJUSFNUQUdFT1VUTEVUVEVNUC5USVQ3OTI4MTE/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 8 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C8_IC4 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQncMCAATkRQSURBQ0" \
                "9MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVA4LkZPVVJUSFNUQUdFSU5URVJDT09MRVJURU1QLlRJVDc5MjgxNQ/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                    print("\033[91mError: Bad transmitter input for Cooler 8 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
                    q_BB2C8_IC4 = 0
            else:
                q_BB2C8_IC4 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # --- COOLER 9 --- 
        volBB2 = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQ3WgDAATkRQ" \
            "SURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9" \
            "OLklOTEVUUkVDSVA5LkNBTENVTEFURURGTE9XLktNNzkyOTIw/value",
            auth=(username, password), verify=False).json() # gets pi data (flowrate and timestamp) for Cooler 9
        if not isinstance(volBB2["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 9 flow...\nDuty marked as 0 BTU/hr...\033[0m")  
                mBB2 = 0
        else:
            mBB2 = volBB2["Value"] * 1/397.3 * mmBB2 / 24 * 1000 # gets mass flow rate [1000-lb/hr]

        # -- INTERCOOLER 1 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQ8GgDAATkRQSURBQ09MTFxVU0" \
            "EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVA5LkZJUlNUU1RBR0VJTkxFVFRFTVAuVElUNzkyOTA4QQ/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 9 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C9_IC1 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQ8WgDAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1R" \
                "BVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVA5LkZJUlNUU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUNzkyOTEy/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 9 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C9_IC1 = 0
            else:
                q_BB2C9_IC1 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 2 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQBGkDAATk" \
            "RQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVA5LlNFQ09ORFNUQUdFT1VUTEVUVEVNUC5USVQ3OTI5MDk/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 9 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C9_IC2 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQA2kDAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQ" \
                "VNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVA5LlNFQ09ORFNUQUdFSU5URVJDT09MRVJURU1QLlRJVDc5MjkxMw/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 9 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C9_IC2 = 0
            else:
                q_BB2C9_IC2 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 3 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQC2kDAATkR" \
            "QSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVA5LlRISVJEU1RBR0VPVVRMRVRURU1QLlRJVDc5MjkxMA/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 9 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C9_IC3 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQCmkDAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEV" \
                "TMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVA5LlRISVJEU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUNzkyOTE0/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                    print("\033[91mError: Bad transmitter input for Cooler 9 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
                    q_BB2C9_IC3 = 0
            else:
                q_BB2C9_IC3 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 4 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQ9GgDAATkRQSURB" \
            "Q09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVA5LkZPVVJUSFNUQUdFT1VUTEVUVEVNUC5USVQ3OTI5MTE/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 9 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C9_IC4 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQ22gDAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTF" \
                "VFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVA5LkZPVVJUSFNUQUdFSU5URVJDT09MRVJURU1QLlRJVDc5MjkxNQ/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_in["Value"], (int, float)):
                    print("\033[91mError: Bad transmitter input for Cooler 9 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
                    q_BB2C9_IC4 = 0
            else:
                q_BB2C9_IC4 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # --- COOLER 10 --- 
        volBB2 = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQ3GgDAATkRQSUR" \
            "BQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOT" \
            "EVUUkVDSVAxMC5DQUxDVUxBVEVERkxPVy5LTTc5MjAyMA/value",
            auth=(username, password), verify=False).json() # gets pi data (flowrate and timestamp) for Cooler 10
        if not isinstance(volBB2["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 10 flow...\nDuty marked as 0 BTU/hr...\033[0m")  
                mBB2 = 0
        else:
            mBB2 = volBB2["Value"] * 1/397.3 * mmBB2 / 24 * 1000 # gets mass flow rate [1000-lb/hr]

        # -- INTERCOOLER 1 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQEmkDAATkRQSURBQ0" \
            "9MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVAxMC5GSVJTVFNUQUdFSU5MRVRURU1QLlRJVDc5MjAwOEE/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 10 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C10_IC1 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQE2kDAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFV" \
                "FQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVAxMC5GSVJTVFNUQUdFSU5URVJDT09MRVJURU1QLlRJVDc5MjAxMg/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                    print("\033[91mError: Bad transmitter input for Cooler 10 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
                    q_BB2C10_IC1 = 0
            else:
                q_BB2C10_IC1 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 2 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQ4mgDAATkRQS" \
            "URBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVAxMC5TRUNPTkRTVEFHRU9VVExFVFRFTVAuVElUNzkyMDA5/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 10 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C10_IC2 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQ4WgDAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5" \
                "HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVAxMC5TRUNPTkRTVEFHRUlOVEVSQ09PTEVSVEVNUC5USVQ3OTIwMTM/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                    print("\033[91mError: Bad transmitter input for Cooler 10 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
                    q_BB2C10_IC2 = 0
            else:
                q_BB2C10_IC2 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 3 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQ6WgDAATkRQSURBQ0" \
            "9MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVAxMC5USElSRFNUQUdFT1VUTEVUVEVNUC5USVQ3OTIwMTA/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 10 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C10_IC3 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQ6GgDAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlV" \
                "UVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVAxMC5USElSRFNUQUdFSU5URVJDT09MRVJURU1QLlRJVDc5MjAxNA/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                    print("\033[91mError: Bad transmitter input for Cooler 10 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
                    q_BB2C10_IC3 = 0
            else:
                q_BB2C10_IC3 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 4 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQFmkDAATkRQ" \
            "SURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQVNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVAxMC5GT1VSVEhTVEFHRU9VVExFVFRFTVAuVElUNzkyMDEx/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)):
                print("\033[91mError: Bad transmitter input for Cooler 10 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BB2C10_IC4 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQ2mgDAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTMi5HQ" \
                "VNQUk9DRVNTLkdBU0NPTVBSRVNTSU9OLklOTEVUUkVDSVAxMC5GT1VSVEhTVEFHRUlOVEVSQ09PTEVSVEVNUC5USVQ3OTIwMTU/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)):
                    print("\033[91mError: Bad transmitter input for Cooler 10 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
                    q_BB2C10_IC4 = 0
            else:
                q_BB2C10_IC4 = mBB2 * cpBB2 * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # ---- END OF BBCS2 COOLERS ----

        break

# BWCS coolers code    
elif station == "BWCS" or station == '2':
    print("You have selected BWCS\n")
    cpBW = 0.5436823 # Cp of inlet gas [BTU/lb/°F] -- NOTE: USING VALUE FOR BBCS2 UNTIL LAB TEST FOR BWCS - 07.09.25
    mmBW = 25.44 # molecular mass of inlet gas [lb/lb-mol] -- NOTE: USING VALUE FOR BBCS2 UNTIL LAB TEST FOR BWCS - 07.09.25

    while True:
        print("\033[94mNow generating cooler analysis of Coolers 1-4...\033[0m\n")

        # --- COOLER 1 --- 
        volBW = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQACcEAATkRQSURBQ09MTFxVU0EuTk" \
            "QuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCVUZGQUxPV0FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDEuQ0FMQ1VMQVRFREZMT1cuS004MjIxMjA/value",
            auth=(username, password), verify=False).json() # gets pi data for flowrate for Cooler 1 [MCFD]
        if not isinstance(volBW["Value"], (int, float)): # if value has bad input
            print("\033[91mError: Bad transmitter input for Cooler 1 flow...\nDuty marked as 0 BTU/hr...\033[0m")  
            mBW = 0
        else:
            mBW = volBW["Value"] * 1/397.3 * mmBW / 24 # gets mass flow rate [1000-lb/hr]
            if volBW["Value"] < 99: # flow rate is in MMSCFD
                mBW = mBW * 1000
            elif volBW["Value"] > 9999: # flow rate is in SCFD
                mBW = mBW / 1000

        # -- INTERCOOLER 1 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQDicEAATkRQSURBQ09MTFxVU" \
            "0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCVUZGQUxPV0FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDEuRklSU1RTVEFHRU9VVExFVFRFTVAuVElUODIyMTA4Qg/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)): # if value has bad input
            print("\033[91mError: Bad transmitter input for Cooler 1 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BWC1_IC1 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQEicEAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT" \
                "01QUkVTU09SU1RBVElPTjpCVUZGQUxPV0FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDEuRklSU1RTVEFHRUlOVEVSQ09PTEVSVEVNUC5USVQ4MjIxMTI/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                print("\033[91mError: Bad transmitter input for Cooler 1 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BWC1_IC1 = 0
            else:
                q_BWC1_IC1 = mBW * cpBW * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 2 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQDycEAATkRQSURBQ09MTFxVU0EuTkQu" \
            "TUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCVUZGQUxPV0FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDEuU0VDT05EU1RBR0VPVVRMRVRURU1QLlRJVDgyMjEwOQ/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)): # if value has bad input
            print("\033[91mError: Bad transmitter input for Cooler 1 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BWC1_IC2 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQEycEAATkRQSURBQ09MTFxVU0EuTk" \
                "QuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCVUZGQUxPV0FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDEuU0VDT05EU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUODIyMTEz/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                print("\033[91mError: Bad transmitter input for Cooler 1 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BWC1_IC2 = 0
            else:
                q_BWC1_IC2 = mBW * cpBW * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 3 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQECcEAATkRQSURBQ09MTFxVU" \
            "0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCVUZGQUxPV0FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDEuVEhJUkRTVEFHRU9VVExFVFRFTVAuVElUODIyMTEw/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)): # if value has bad input
            print("\033[91mError: Bad transmitter input for Cooler 1 IC-3 ...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BWC1_IC3 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQFCcEAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCVU" \
                "ZGQUxPV0FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDEuVEhJUkRTVEFHRUlOVEVSQ09PTEVSVEVNUC5USVQ4MjIxMTQ/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                print("\033[91mError: Bad transmitter input for Cooler 1 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BWC1_IC3 = 0
            else:
                q_BWC1_IC3 = mBW * cpBW * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 4 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQEScEAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT0" \
            "1QUkVTU09SU1RBVElPTjpCVUZGQUxPV0FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDEuRk9VUlRIU1RBR0VPVVRMRVRURU1QLlRJVDgyMjExMQ/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)): # if value has bad input
            print("\033[91mError: Bad transmitter input for Cooler 1 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BWC1_IC4 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQFScEAATkRQSURBQ09MTFxVU0" \
                "EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCVUZGQUxPV0FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDEuRk9VUlRIU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUODIyMTE1/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                print("\033[91mError: Bad transmitter input for Cooler 1 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BWC1_IC4 = 0
            else:
                q_BWC1_IC4 = mBW * cpBW * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # --- COOLER 2 --- 
        volBW = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQMScEAATkRQSURBQ09MTFx" \
            "VU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCVUZGQUxPV0FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDIuQ0FMQ1VMQVRFREZMT1cuS004MjIyMjA/value",
            auth=(username, password), verify=False).json() # gets pi data for flowrate for Cooler 2 [MCFD]
        if not isinstance(volBW["Value"], (int, float)): # if value has bad input
            print("\033[91mError: Bad transmitter input for Cooler 2 flow...\nDuty marked as 0 BTU/hr...\033[0m")  
            mBW = 0
        else:
            mBW = volBW["Value"] * 1/397.3 * mmBW / 24 # gets mass flow rate [1000-lb/hr]
            if volBW["Value"] < 99: # flow rate is in MMSCFD
                mBW = mBW * 1000
            elif volBW["Value"] > 9999: # flow rate is in SCFD
                mBW = mBW / 1000

        # -- INTERCOOLER 1 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQPycEAATkRQSURBQ09MT" \
            "FxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCVUZGQUxPV0FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDIuRklSU1RTVEFHRU9VVExFVFRFTVAuVElUODIyMjA4Qg/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)): # if value has bad input
            print("\033[91mError: Bad transmitter input for Cooler 2 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BWC2_IC1 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQQycEAATkRQSURBQ09MTFxVU0EuTkQuTUlELl" \
                "NHUy5DT01QUkVTU09SU1RBVElPTjpCVUZGQUxPV0FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDIuRklSU1RTVEFHRUlOVEVSQ09PTEVSVEVNUC5USVQ4MjIyMTI/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                print("\033[91mError: Bad transmitter input for Cooler 2 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BWC2_IC1 = 0
            else:
                q_BWC2_IC1 = mBW * cpBW * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 2 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQQCcEAATkRQSURBQ09MTFx" \
            "VU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCVUZGQUxPV0FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDIuU0VDT05EU1RBR0VPVVRMRVRURU1QLlRJVDgyMjIwOQ/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)): # if value has bad input
            print("\033[91mError: Bad transmitter input for Cooler 2 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BWC2_IC2 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQRCcEAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT" \
                "01QUkVTU09SU1RBVElPTjpCVUZGQUxPV0FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDIuU0VDT05EU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUODIyMjEz/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                print("\033[91mError: Bad transmitter input for Cooler 2 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BWC2_IC2 = 0
            else:
                q_BWC2_IC2 = mBW * cpBW * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 3 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQQScEAATkRQSURBQ09MTFxVU0EuTkQuTUlELlN" \
            "HUy5DT01QUkVTU09SU1RBVElPTjpCVUZGQUxPV0FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDIuVEhJUkRTVEFHRU9VVExFVFRFTVAuVElUODIyMjEw/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)): # if value has bad input
            print("\033[91mError: Bad transmitter input for Cooler 2 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BWC2_IC3 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQRScEAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09" \
                "SU1RBVElPTjpCVUZGQUxPV0FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDIuVEhJUkRTVEFHRUlOVEVSQ09PTEVSVEVNUC5USVQ4MjIyMTQ/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                print("\033[91mError: Bad transmitter input for Cooler 2 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BWC2_IC3 = 0
            else:
                q_BWC2_IC3 = mBW * cpBW * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 4 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQQicEAATkRQSURBQ09MTFxVU0" \
            "EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCVUZGQUxPV0FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDIuRk9VUlRIU1RBR0VPVVRMRVRURU1QLlRJVDgyMjIxMQ/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)): # if value has bad input
            print("\033[91mError: Bad transmitter input for Cooler 2 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BWC2_IC4 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQRicEAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1R" \
                "VElPTjpCVUZGQUxPV0FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDIuRk9VUlRIU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUODIyMjE1/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                print("\033[91mError: Bad transmitter input for Cooler 2 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BWC2_IC4 = 0
            else:
                q_BWC2_IC4 = mBW * cpBW * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # --- COOLER 3 --- 
        volBW = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQYycEAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCVUZGQUxPV0" \
            "FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDMuQ0FMQ1VMQVRFREZMT1cuS004MjIzMjA/value",
            auth=(username, password), verify=False).json() # gets pi data for flowrate for Cooler 3 [MCFD]
        if not isinstance(volBW["Value"], (int, float)): # if value has bad input
            print("\033[91mError: Bad transmitter input for Cooler 3 flow...\nDuty marked as 0 BTU/hr...\033[0m")  
            mBW = 0
        else:
            mBW = volBW["Value"] * 1/397.3 * mmBW / 24 # gets mass flow rate [1000-lb/hr]
            if volBW["Value"] < 99: # flow rate is in MMSCFD
                mBW = mBW * 1000
            elif volBW["Value"] > 9999: # flow rate is in SCFD
                mBW = mBW / 1000

        # -- INTERCOOLER 1 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQcScEAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5" \
            "DT01QUkVTU09SU1RBVElPTjpCVUZGQUxPV0FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDMuRklSU1RTVEFHRU9VVExFVFRFTVAuVElUODIyMzA4Qg/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)): # if value has bad input
            print("\033[91mError: Bad transmitter input for Cooler 3 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BWC3_IC1 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQdScEAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBV" \
                "ElPTjpCVUZGQUxPV0FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDMuRklSU1RTVEFHRUlOVEVSQ09PTEVSVEVNUC5USVQ4MjIzMTI/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                print("\033[91mError: Bad transmitter input for Cooler 3 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BWC3_IC1 = 0
            else:
                q_BWC3_IC1 = mBW * cpBW * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 2 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQdicEAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCVUZGQUxPV0FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDMuU" \
            "0VDT05EU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUODIyMzEz/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)): # if value has bad input
            print("\033[91mError: Bad transmitter input for Cooler 3 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BWC3_IC2 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQdicEAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09" \
                "SU1RBVElPTjpCVUZGQUxPV0FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDMuU0VDT05EU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUODIyMzEz/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                print("\033[91mError: Bad transmitter input for Cooler 3 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BWC3_IC2 = 0
            else:
                q_BWC3_IC2 = mBW * cpBW * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 3 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQcycEAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1R" \
            "BVElPTjpCVUZGQUxPV0FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDMuVEhJUkRTVEFHRU9VVExFVFRFTVAuVElUODIyMzEw/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)): # if value has bad input
            print("\033[91mError: Bad transmitter input for Cooler 3 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BWC3_IC3 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQdycEAATkRQSURBQ09MTFxVU0EuTkQuTUlELl" \
                "NHUy5DT01QUkVTU09SU1RBVElPTjpCVUZGQUxPV0FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDMuVEhJUkRTVEFHRUlOVEVSQ09PTEVSVEVNUC5USVQ4MjIzMTQ/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                print("\033[91mError: Bad transmitter input for Cooler 3 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BWC3_IC3 = 0
            else:
                q_BWC3_IC3 = mBW * cpBW * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 4 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQdCcEAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01Q" \
            "UkVTU09SU1RBVElPTjpCVUZGQUxPV0FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDMuRk9VUlRIU1RBR0VPVVRMRVRURU1QLlRJVDgyMjMxMQ/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)): # if value has bad input
            print("\033[91mError: Bad transmitter input for Cooler 3 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BWC3_IC4 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQeCcEAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUk" \
                "VTU09SU1RBVElPTjpCVUZGQUxPV0FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDMuRk9VUlRIU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUODIyMzE1/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                print("\033[91mError: Bad transmitter input for Cooler 3 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BWC3_IC4 = 0
            else:
                q_BWC3_IC4 = mBW * cpBW * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # --- COOLER 4 --- 
        volBW = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQlScEAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElP" \
            "TjpCVUZGQUxPV0FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDQuQ0FMQ1VMQVRFREZMT1cuS004MjI0MjA/value",
            auth=(username, password), verify=False).json() # gets pi data for flowrate for Cooler 4 [MCFD]
        if not isinstance(volBW["Value"], (int, float)): # if value has bad input
            print("\033[91mError: Bad transmitter input for Cooler 4 flow...\nDuty marked as 0 BTU/hr...\033[0m")  
            mBW = 0
        else:
            mBW = volBW["Value"] * 1/397.3 * mmBW / 24 # gets mass flow rate [1000-lb/hr]
            if volBW["Value"] < 99: # flow rate is in MMSCFD
                mBW = mBW * 1000
            elif volBW["Value"] > 9999: # flow rate is in SCFD
                mBW = mBW / 1000

        # -- INTERCOOLER 1 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQoycEAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01" \
            "QUkVTU09SU1RBVElPTjpCVUZGQUxPV0FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDQuRklSU1RTVEFHRU9VVExFVFRFTVAuVElUODIyNDA4Qg/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)): # if value has bad input
            print("\033[91mError: Bad transmitter input for Cooler 4 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BWC4_IC1 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQpycEAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RB" \
                "VElPTjpCVUZGQUxPV0FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDQuRklSU1RTVEFHRUlOVEVSQ09PTEVSVEVNUC5USVQ4MjI0MTI/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                print("\033[91mError: Bad transmitter input for Cooler 4 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BWC4_IC1 = 0
            else:
                q_BWC4_IC1 = mBW * cpBW * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 2 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQpCcEAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVT" \
            "U09SU1RBVElPTjpCVUZGQUxPV0FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDQuU0VDT05EU1RBR0VPVVRMRVRURU1QLlRJVDgyMjQwOQ/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)): # if value has bad input
            print("\033[91mError: Bad transmitter input for Cooler 4 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BWC4_IC2 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQqCcEAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVE" \
                "lPTjpCVUZGQUxPV0FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDQuU0VDT05EU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUODIyNDEz/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                print("\033[91mError: Bad transmitter input for Cooler 4 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BWC4_IC2 = 0
            else:
                q_BWC4_IC2 = mBW * cpBW * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 3 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQpScEAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU" \
            "09SU1RBVElPTjpCVUZGQUxPV0FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDQuVEhJUkRTVEFHRU9VVExFVFRFTVAuVElUODIyNDEw/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)): # if value has bad input
            print("\033[91mError: Bad transmitter input for Cooler 4 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BWC4_IC3 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQqScEAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1R" \
                "BVElPTjpCVUZGQUxPV0FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDQuVEhJUkRTVEFHRUlOVEVSQ09PTEVSVEVNUC5USVQ4MjI0MTQ/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                print("\033[91mError: Bad transmitter input for Cooler 4 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BWC4_IC3 = 0
            else:
                q_BWC4_IC3 = mBW * cpBW * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 4 --
        # gaining values for q computation
        T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQpicEAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVE" \
            "lPTjpCVUZGQUxPV0FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDQuRk9VUlRIU1RBR0VPVVRMRVRURU1QLlRJVDgyMjQxMQ/value",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)): # if value has bad input
            print("\033[91mError: Bad transmitter input for Cooler 4 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BWC4_IC4 = 0
        else:
            T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQqicEAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCVU" \
                "GQUxPV0FMTE9XLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDQuRk9VUlRIU1RBR0VJTlRFUkNPT0xFUlRFTVAuVElUODIyNDE1/value",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                print("\033[91mError: Bad transmitter input for Cooler 4 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BWC4_IC4 = 0
            else:
                q_BWC4_IC4 = mBW * cpBW * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)
        
        '''
        NOTE: ADDED CODE FOR ADDITIONAL COOLERS WHEN STATION EXPANDED. REMOVE COMMENT 'TRIPLE APOSTROPHE' AS NEEDED TO UNCOMMENT

        # --- COOLER 5 --- 
        volBW = requests.get("",
            auth=(username, password), verify=False).json() # gets pi data for flowrate for Cooler 5 [MCFD]
        if not isinstance(volBW["Value"], (int, float)): # if value has bad input
            print("\033[91mError: Bad transmitter input for Cooler 5 flow...\nDuty marked as 0 BTU/hr...\033[0m")  
            mBW = 0
        else:
            mBW = volBW["Value"] * 1/397.3 * mmBW / 24 # gets mass flow rate [1000-lb/hr]
            if volBW["Value"] < 99: # flow rate is in MMSCFD
                mBW = mBW * 1000
            elif volBW["Value"] > 9999: # flow rate is in SCFD
                mBW = mBW / 1000

        # -- INTERCOOLER 1 --
        # gaining values for q computation
        T_in = requests.get("",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)): # if value has bad input
            print("\033[91mError: Bad transmitter input for Cooler 5 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BWC5_IC1 = 0
        else:
            T_out = requests.get("",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                print("\033[91mError: Bad transmitter input for Cooler 5 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BWC5_IC1 = 0
            else:
                q_BWC5_IC1 = mBW * cpBW * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)
        
        # -- INTERCOOLER 2 --
        # gaining values for q computation
        T_in = requests.get("",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)): # if value has bad input
            print("\033[91mError: Bad transmitter input for Cooler 5 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BWC5_IC2 = 0
        else:
            T_out = requests.get("",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                print("\033[91mError: Bad transmitter input for Cooler 5 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BWC5_IC2 = 0
            else:
                q_BWC5_IC2 = mBW * cpBW * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 3 --
        # gaining values for q computation
        T_in = requests.get("",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)): # if value has bad input
            print("\033[91mError: Bad transmitter input for Cooler 5 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BWC5_IC3 = 0
        else:
            T_out = requests.get("",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                print("\033[91mError: Bad transmitter input for Cooler 5 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BWC5_IC3 = 0
            else:
                q_BWC5_IC3 = mBW * cpBW * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 4 --
        # gaining values for q computation
        T_in = requests.get("",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)): # if value has bad input
            print("\033[91mError: Bad transmitter input for Cooler 5 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BWC5_IC4 = 0
        else:
            T_out = requests.get("",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                print("\033[91mError: Bad transmitter input for Cooler 5 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BWC5_IC4 = 0
            else:
                q_BWC5_IC4 = mBW * cpBW * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # --- COOLER 6 --- 
        volBW = requests.get("",
            auth=(username, password), verify=False).json() # gets pi data for flowrate for Cooler 6 [MCFD]
        if not isinstance(volBW["Value"], (int, float)): # if value has bad input
            print("\033[91mError: Bad transmitter input for Cooler 6 flow...\nDuty marked as 0 BTU/hr...\033[0m")  
            mBW = 0
        else:
            mBW = volBW["Value"] * 1/397.3 * mmBW / 24 # gets mass flow rate [1000-lb/hr]
            if volBW["Value"] < 99: # flow rate is in MMSCFD
                mBW = mBW * 1000
            elif volBW["Value"] > 9999: # flow rate is in SCFD
                mBW = mBW / 1000

        # -- INTERCOOLER 1 --
        # gaining values for q computation
        T_in = requests.get("",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)): # if value has bad input
            print("\033[91mError: Bad transmitter input for Cooler 6 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BWC6_IC1 = 0
        else:
            T_out = requests.get("",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                print("\033[91mError: Bad transmitter input for Cooler 6 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BWC6_IC1 = 0
            else:
                q_BWC6_IC1 = mBW * cpBW * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)
        
        # -- INTERCOOLER 2 --
        # gaining values for q computation
        T_in = requests.get("",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)): # if value has bad input
            print("\033[91mError: Bad transmitter input for Cooler 6 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BWC6_IC2 = 0
        else:
            T_out = requests.get("",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                print("\033[91mError: Bad transmitter input for Cooler 6 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BWC6_IC2 = 0
            else:
                q_BWC6_IC2 = mBW * cpBW * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 3 --
        # gaining values for q computation
        T_in = requests.get("",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)): # if value has bad input
            print("\033[91mError: Bad transmitter input for Cooler 6 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BWC6_IC3 = 0
        else:
            T_out = requests.get("",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                print("\033[91mError: Bad transmitter input for Cooler 6 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BWC6_IC3 = 0
            else:
                q_BWC6_IC3 = mBW * cpBW * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

        # -- INTERCOOLER 4 --
        # gaining values for q computation
        T_in = requests.get("",
            auth=(username, password), verify=False).json() # gets pi data for inlet temp
        if not isinstance(T_in["Value"], (int, float)): # if value has bad input
            print("\033[91mError: Bad transmitter input for Cooler 6 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
            q_BWC6_IC4 = 0
        else:
            T_out = requests.get("",
                auth=(username, password), verify=False).json() # gets pi data for outlet temp
            if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                print("\033[91mError: Bad transmitter input for Cooler 6 IC-4...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BWC6_IC4 = 0
            else:
                q_BWC6_IC4 = mBW * cpBW * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)
                
        '''

        # --- END OF BWCS COOLERS ---

        break

# BBCS legacy cooler 4 code (ASSUMING COMP4 ONLY RUNNING COMPRESSOR) 
elif station == "BBCS" or station == '3':
    print("You have selected BBCS Compressor 4\n")
    cpBB = 0.5436823 # Cp of inlet gas [BTU/lb/°F]
    mmBB = 25.44 # molecular mass of inlet gas [lb/lb-mol]

    while True:
        input("Verify that the operator rounds data is uploaded to this folder with the name 'BBCSRounds.xlsx' and press ENTER to continue...\n")
        T_out_list = pd.read_excel("BBCSRounds.xlsx")[['Label', 'Response']]
        status = (T_out_list[T_out_list['Label'] == 'COMPRESSOR STATUS'])['Response'].tolist()
        print("\033[94mNow entering temps for Cooler 4...\033[0m\n")
        if status[3].strip().upper() == 'OOS' or status[3].strip().upper() == 'STANDBY':
                print("\033[91mError: Comp 4 OOS/STANDBY...\nUnable to compute duties. Exiting program...\033[0m")  
                sys.exit()
        else:
            # --- COOLER 4 --- 
            volBB = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQxvIAAATkRQSURBQ" \
                "09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTTEVHLlNBTEVTJlBSRFRSTlNQLkdBU01FVEVSSU5HLkdBU01FVEVSSU5HLlNUQVRJT05ESVNDSEFSR0VHQVNGTE9XWURZLkZRSVQ1MjAwMDE/value",
                auth=(username, password), verify=False).json() # gets pi data (flowrate and timestamp) for Cooler 4
            if not isinstance(volBB["Value"], (int, float)): # if value has bad input
                print("\033[91mError: Bad transmitter input for Cooler 4 flow...\nDuty marked as 0 BTU/hr...\033[0m")  
                volBB = 0
            else:
                mBB = volBB["Value"] * 1/397.3 * mmBB / 24 # gets mass flow rate [1000-lb/hr]

            T_out_list['Response'] = pd.to_numeric((T_out_list['Response'].astype(str).str.replace(' degF', '', regex=False)), errors='coerce')

            # -- INTERCOOLER 1 --
            # gaining values for q computation
            T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQuPIAAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTT" \
                "EVHLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDQuRklSU1RTVEFHRU9VVExFVFRFTVAuVElUNTIyNDAy/value",
                auth=(username, password), verify=False).json() # gets pi data for inlet temp
            if not isinstance(T_in["Value"], (int, float)): # if value has bad input
                print("\033[91mError: Bad transmitter input for Cooler 4 IC-1...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BBC4_IC1 = 0
            else:
                T_out = (T_out_list[T_out_list['Label'] == '1ST STAGE INTER TEMP'].iloc[0])['Response']
                q_BBC4_IC1 = mBB * cpBB * (T_in["Value"] - T_out) # heat rejection (1000-BTU/hr)

            # -- INTERCOOLER 2 --
            # gaining values for q computation
            T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQu_IAAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTTEVHL" \
                "kdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDQuU0VDT05EU1RBR0VPVVRMRVRURU1QLlRJVDUyMjQwNg/value",
                auth=(username, password), verify=False).json() # gets pi data for inlet temp
            if not isinstance(T_in["Value"], (int, float)): # if value has bad input
                print("\033[91mError: Bad transmitter input for Cooler 4 IC-2...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BBC4_IC2 = 0
            else:
                T_out = (T_out_list[T_out_list['Label'] == '2ND STAGE INTER TEMP'].iloc[0])['Response']
                q_BBC4_IC2 = mBB * cpBB * (T_in["Value"] - T_out) # heat rejection (1000-BTU/hr)

            # -- INTERCOOLER 3 --
            # gaining values for q computation
            T_in = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQvvIAAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTTEV" \
                "HLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDQuVEhJUkRTVEFHRU9VVExFVFRFTVAuVElUNTIyNDA3/value",
                auth=(username, password), verify=False).json() # gets pi data for inlet temp
            if not isinstance(T_in["Value"], (int, float)): # if value has bad input
                print("\033[91mError: Bad transmitter input for Cooler 4 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
                q_BBC4_IC3 = 0
            else:
                T_out = requests.get("https://ndpi.vision.ihess.com/piwebapi/streams/F1DPvVzCIaSjSkCZuCntY8tYTQrvIAAATkRQSURBQ09MTFxVU0EuTkQuTUlELlNHUy5DT01QUkVTU09SU1RBVElPTjpCTFVFQlVUVEVTT" \
                    "EVHLkdBU1BST0NFU1MuR0FTQ09NUFJFU1NJT04uSU5MRVRSRUNJUDQuRklOQUxESVNDSEFSR0VURU1QLlRJVDUyMjQwNA/value",
                    auth=(username, password), verify=False).json() # gets pi data for outlet temp
                if not isinstance(T_out["Value"], (int, float)): # if value has bad input
                    print("\033[91mError: Bad transmitter input for Cooler 4 IC-3...\nDuty marked as 0 BTU/hr...\033[0m")  
                    q_BBC4_IC3 = 0
                else:
                    q_BBC4_IC3 = mBB * cpBB * (T_in["Value"] - T_out["Value"]) # heat rejection (1000-BTU/hr)

            break

else:
    print("\n\033[91mInvalid compressor station. Try again.\033[0m\n")
    sys.exit()

# AFTER COLLECTING DATA

print("\n\033[92mAnalysis generation complete\033[0m")

while True:
    save = input("\nConfirm that you would like to save findings (Y/N).\n").strip().lower()
    if save == 'y': # if user selects to keep readings
        timestamp = datetime.today()

        # bbcs2 plotting
        if station == "BBCS2" or station == '1':
            BBCS2df = pd.DataFrame({
            "Timestamp": timestamp,
            "Cooler 1 IC Duties [1000-BTU/hr]": [q_BB2C1_IC1, q_BB2C1_IC2, q_BB2C1_IC3, q_BB2C1_IC4],
            "Cooler 2 IC Duties [1000-BTU/hr]": [q_BB2C2_IC1, q_BB2C2_IC2, q_BB2C2_IC3, q_BB2C2_IC4],
            "Cooler 3 IC Duties [1000-BTU/hr]": [q_BB2C3_IC1, q_BB2C3_IC2, q_BB2C3_IC3, q_BB2C3_IC4],
            "Cooler 4 IC Duties [1000-BTU/hr]": [q_BB2C4_IC1, q_BB2C4_IC2, q_BB2C4_IC3, q_BB2C4_IC4],
            "Cooler 5 IC Duties [1000-BTU/hr]": [q_BB2C5_IC1, q_BB2C5_IC2, q_BB2C5_IC3, q_BB2C5_IC4],
            "Cooler 6 IC Duties [1000-BTU/hr]": [q_BB2C6_IC1, q_BB2C6_IC2, q_BB2C6_IC3, q_BB2C6_IC4],
            "Cooler 7 IC Duties [1000-BTU/hr]": [q_BB2C7_IC1, q_BB2C7_IC2, q_BB2C7_IC3, q_BB2C7_IC4],
            "Cooler 8 IC Duties [1000-BTU/hr]": [q_BB2C8_IC1, q_BB2C8_IC2, q_BB2C8_IC3, q_BB2C8_IC4],
            "Cooler 9 IC Duties [1000-BTU/hr]": [q_BB2C9_IC1, q_BB2C9_IC2, q_BB2C9_IC3, q_BB2C9_IC4],
            "Cooler 10 IC Duties [1000-BTU/hr]": [q_BB2C10_IC1, q_BB2C10_IC2, q_BB2C10_IC3, q_BB2C10_IC4]
            })
            
            # writing data to excel sheet  
            sheetName = datetime.now().strftime("Run_%m%d%Y_%H%M%S") # labels sheet

            workbook = load_workbook("BBCS2Coolers.xlsx")
            writer = pd.ExcelWriter("BBCS2Coolers.xlsx", engine="openpyxl", mode="a")

            BBCS2df.to_excel(writer, sheet_name=sheetName, index=False)

            writer.close()

            print("\033[92m'BBCS2Coolers.xlsx' now updated\033[0m\n")

            print("\n\033[92mNow generating plots...\033[0m")

            # generating plots for coolers
            # Load Excel sheets
            xls = pd.read_excel("BBCS2Coolers.xlsx", sheet_name=None, engine="openpyxl")
            existing_df = next(iter(xls.values()))
            time_col = existing_df.columns[0]

            # Convert time column to datetime
            for name, df in xls.items():
                df[time_col] = pd.to_datetime(df[time_col])

            cooler_cols = [col for col in existing_df.columns if col != time_col]

            for col_index, col in enumerate(cooler_cols):
                    plt.figure(figsize=(8, 5))
                    scatter_objects = []

                    for sheet_name, df in xls.items():
                        x_vals = df[time_col]
                        y_vals = df[col]

                        # Scatter plot
                        sc = plt.scatter(x_vals, y_vals, label=sheet_name)
                        scatter_objects.append((sc, df, sheet_name, col))

                        labels = ['IC-1', 'IC-2', 'IC-3', 'IC-4']

                        # Optional: annotate points if needed
                        for i in range(len(x_vals)):
                            plt.annotate(
                                 labels[i],
                                 (x_vals.iloc[i], y_vals.iloc[i]),
                                 textcoords="offset points",
                                 xytext=(5, 5),
                                 ha='center',
                                 fontsize=8,
                                 color='gray'
                            )

                    # Define dashed lines once per figure
                    dashed_lines = {
                        'IC-1 Duty': 935.2,
                        'IC-2 Duty': 1663.6,
                        'IC-3 Duty': 1831.6,
                        'IC-4 Duty': 2187.3,
                    }

                    # Draw dashed lines and label once
                    for label, y_val in dashed_lines.items():
                        plt.axhline(y=y_val, color='black', linestyle='--')
                        plt.text(
                            plt.gca().get_xlim()[0],
                            y_val + 12,
                            label,
                            color='black',
                            fontsize=7,
                            ha='left',
                            va='bottom'
                        )

                    plt.title(f"{col} vs {time_col}")
                    plt.xlabel(time_col)
                    plt.ylabel(col)
                    plt.grid(True)
                    plt.tight_layout()


                    # Add interactivity for this figure
                    for sc, df, sheet_name, col in scatter_objects:
                        cursor = mplcursors.cursor(sc, hover=True)
                        @cursor.connect("add")
                        def on_add(sel, df=df, sheet_name=sheet_name, col=col):
                            i = sel.index
                            sel.annotation.set_text(
                                f"{sheet_name}\n{labels[i]}\n{df[time_col].iloc[i]}\n{df[col].iloc[i]:.2f}"
                            )

            plt.show()
            
            break

        # hgf plotting
        elif station == 'HGF' or station == '0':
            if cooler == 'inlet gas' or cooler == '0':
                HGFdf = pd.DataFrame({
                "Timestamp": timestamp,
                "Cooler 1 IC Duties [1000-BTU/hr]": [q_C1_IC1, q_C1_IC2, q_C1_IC3, q_C1_IC4, q_C1_EJW],
                "Cooler 2 IC Duties [1000-BTU/hr]": [q_C2_IC1, q_C2_IC2, q_C2_IC3, q_C2_IC4, q_C2_EJW],
                "Cooler 3 IC Duties [1000-BTU/hr]": [q_C3_IC1, q_C3_IC2, q_C3_IC3, q_C3_IC4, q_C3_EJW],
                "Cooler 4 IC Duties [1000-BTU/hr]": [q_C4_IC1, q_C4_IC2, q_C4_IC3, q_C4_IC4, q_C4_EJW],
                "Cooler 5 IC Duties [1000-BTU/hr]": [q_C5_IC1, q_C5_IC2, q_C5_IC3, q_C5_IC4, q_C5_EJW]
                })
                
                # writing data to excel sheet  
                sheetName = datetime.now().strftime("Run_%m%d%Y_%H%M%S") # labels sheet

                workbook = load_workbook("HGFInletCoolers.xlsx")
                writer = pd.ExcelWriter("HGFInletCoolers.xlsx", engine="openpyxl", mode="a")

                HGFdf.to_excel(writer, sheet_name=sheetName, index=False)

                writer.close()

                print("\033[92m'HGFInletCoolers.xlsx' now updated\033[0m")

                print("\n\033[92mNow generating plots...\033[0m")

                # generating plots for coolers
                # Load Excel sheets
                xls = pd.read_excel("HGFInletCoolers.xlsx", sheet_name=None, engine="openpyxl")
                existing_df = next(iter(xls.values()))
                time_col = existing_df.columns[0]

                # Convert time column to datetime
                for name, df in xls.items():
                    df[time_col] = pd.to_datetime(df[time_col])

                cooler_cols = [col for col in existing_df.columns if col != time_col]

                for col_index, col in enumerate(cooler_cols):
                    plt.figure(figsize=(8, 5))
                    scatter_objects = []

                    for sheet_name, df in xls.items():
                        x_vals = df[time_col]
                        y_vals = df[col]

                        # Scatter plot
                        sc = plt.scatter(x_vals, y_vals, label=sheet_name)
                        scatter_objects.append((sc, df, sheet_name, col))

                        labels = ['IC-1', 'IC-2', 'IC-3', 'IC-4', 'EJW']

                        # Optional: annotate points if needed
                        for i in range(len(x_vals)):
                            plt.annotate(
                                 labels[i],
                                 (x_vals.iloc[i], y_vals.iloc[i]),
                                 textcoords="offset points",
                                 xytext=(5, 5),
                                 ha='center',
                                 fontsize=8,
                                 color='gray'
                            )

                    # Define dashed lines once per figure
                    if col_index < 3:
                        dashed_lines = {
                            'IC-1 Duty': 3050,
                            'IC-2 Duty': 2516,
                            'IC-3 Duty': 2958,
                            'IC-4 Duty': 4103,
                            'EJW Duty': 2654
                        }
                    else:
                        dashed_lines = {
                            'IC-1 Duty': 1050.1,
                            'IC-2 Duty': 1270.6,
                            'IC-3 Duty': 1561,
                            'IC-4 Duty': 2433.5,
                            'EJW Duty': 1305.8
                        }

                    # Draw dashed lines and label once
                    for label, y_val in dashed_lines.items():
                        plt.axhline(y=y_val, color='black', linestyle='--')
                        plt.text(
                            plt.gca().get_xlim()[0],
                            y_val + 12,
                            label,
                            color='black',
                            fontsize=7,
                            ha='left',
                            va='bottom'
                        )

                    plt.title(f"{col} vs {time_col}")
                    plt.xlabel(time_col)
                    plt.ylabel(col)
                    plt.grid(True)
                    plt.tight_layout()


                    # Add interactivity for this figure
                    for sc, df, sheet_name, col in scatter_objects:
                        cursor = mplcursors.cursor(sc, hover=True)
                        @cursor.connect("add")
                        def on_add(sel, df=df, sheet_name=sheet_name, col=col):
                            i = sel.index
                            sel.annotation.set_text(
                                f"{sheet_name}\n{labels[i]}\n{df[time_col].iloc[i]}\n{df[col].iloc[i]:.2f}"
                            )

                plt.show()
            elif cooler == 'flash gas' or cooler == '1':
                HGFdf = pd.DataFrame({
                "Timestamp": timestamp,
                "Cooler 1 IC Duties [1000-BTU/hr]": [q_FG1_IC1],
                # "Cooler 2 IC Duties [1000-BTU/hr]": [q_FG2_IC1] # COMMENTED OUT WHILE NOT OOS
                # "Cooler 3 IC Duties [1000-BTU/hr]": q_FG3_IC1 # COMMENTED OUT WHILE NOT OOS
                })
                
                # writing data to excel sheet  
                sheetName = datetime.now().strftime("Run_%m%d%Y_%H%M%S") # labels sheet

                workbook = load_workbook("HGFFGCoolers.xlsx")
                writer = pd.ExcelWriter("HGFFGCoolers.xlsx", engine="openpyxl", mode="a")

                HGFdf.to_excel(writer, sheet_name=sheetName, index=False)

                writer.close()

                print("\033[92m'HGFFGCoolers.xlsx' now updated\033[0m")

                print("\n\033[92mNow generating plots...\033[0m")

                # generating plots for coolers
                # Load Excel sheets
                xls = pd.read_excel("HGFFGCoolers.xlsx", sheet_name=None, engine="openpyxl")
                existing_df = next(iter(xls.values()))
                time_col = existing_df.columns[0]

                # Convert time column to datetime
                for name, df in xls.items():
                    df[time_col] = pd.to_datetime(df[time_col])

                cooler_cols = [col for col in existing_df.columns if col != time_col]

                for col_index, col in enumerate(cooler_cols):
                    plt.figure(figsize=(8, 5))
                    scatter_objects = []

                    for sheet_name, df in xls.items():
                        x_vals = df[time_col]
                        y_vals = df[col]

                        # Scatter plot
                        sc = plt.scatter(x_vals, y_vals, label=sheet_name)
                        scatter_objects.append((sc, df, sheet_name, col))

                        labels = ['IC-1', 'EJW']

                        # Optional: annotate points if needed
                        for i in range(len(x_vals)):
                            plt.annotate(
                                 labels[i],
                                 (x_vals.iloc[i], y_vals.iloc[i]),
                                 textcoords="offset points",
                                 xytext=(5, 5),
                                 ha='center',
                                 fontsize=8,
                                 color='gray'
                            )

                    # Define dashed lines once per figure
                    dashed_lines = {
                        'IC-1 Duty': 2431.6,
                        'EJW Duty': 1448.4
                    }

                    # Draw dashed lines and label once
                    for label, y_val in dashed_lines.items():
                        plt.axhline(y=y_val, color='black', linestyle='--')
                        plt.text(
                            plt.gca().get_xlim()[0],
                            y_val + 12,
                            label,
                            color='black',
                            fontsize=7,
                            ha='left',
                            va='bottom'
                        )

                    plt.title(f"{col} vs {time_col}")
                    plt.xlabel(time_col)
                    plt.ylabel(col)
                    plt.grid(True)
                    plt.tight_layout()

                    # Add interactivity for this figure
                    for sc, df, sheet_name, col in scatter_objects:
                        cursor = mplcursors.cursor(sc, hover=True)
                        @cursor.connect("add")
                        def on_add(sel, df=df, sheet_name=sheet_name, col=col):
                            i = sel.index
                            sel.annotation.set_text(
                                f"{sheet_name}\n{labels[i]}\n{df[time_col].iloc[i]}\n{df[col].iloc[i]:.2f}"
                            )

                plt.show()
            elif cooler == 'refrig' or cooler == '2':
                HGFdf = pd.DataFrame({
                "Timestamp": timestamp,
                "Cooler 1 EJW Duty [1000-BTU/hr]": [q_Re1_EJW],
                #"Cooler 2 EJW Duty [1000-BTU/hr]": [q_Re2_EJW] # COMMENTED OUT WHILE ONLY 1 UNIT ACTIEV
                })
                
                # writing data to excel sheet  
                sheetName = datetime.now().strftime("Run_%m%d%Y_%H%M%S") # labels sheet

                workbook = load_workbook("HGFRefrigCoolers.xlsx")
                writer = pd.ExcelWriter("HGFRefrigCoolers.xlsx", engine="openpyxl", mode="a")

                HGFdf.to_excel(writer, sheet_name=sheetName, index=False)

                writer.close()

                print("\033[92m'HGFRefrigCoolers.xlsx' now updated\033[0m")

                print("\n\033[92mNow generating plots...\033[0m")

                # generating plots for coolers
                # Load Excel sheets
                xls = pd.read_excel("HGFRefrigCoolers.xlsx", sheet_name=None, engine="openpyxl")
                existing_df = next(iter(xls.values()))
                time_col = existing_df.columns[0]

                # Convert time column to datetime
                for name, df in xls.items():
                    df[time_col] = pd.to_datetime(df[time_col])

                cooler_cols = [col for col in existing_df.columns if col != time_col]

                for col in cooler_cols:
                    plt.figure(figsize=(8, 5))
                    scatter_objects = []  # Store scatter plots for this figure

                    for sheet_name, df in xls.items():
                        x_vals = df[time_col]
                        y_vals = df[col]

                        # Scatter plot
                        sc = plt.scatter(x_vals, y_vals, label=sheet_name)
                        scatter_objects.append((sc, df, sheet_name, col))

                        labels = ["EJW"]

                        # Static labels
                        for i in range(len(x_vals)):
                            plt.annotate(
                                labels[i],
                                (x_vals.iloc[i], y_vals.iloc[i]),
                                textcoords="offset points",
                                xytext=(5, 5),
                                ha='center',
                                fontsize=8,
                                color='gray'
                            )

                    plt.title(f"{col} vs {time_col}")
                    plt.xlabel(time_col)
                    plt.ylabel(col)
                    plt.legend()
                    plt.grid(True)
                    plt.tight_layout()

                    # Add interactivity for this figure
                    for sc, df, sheet_name, col in scatter_objects:
                        cursor = mplcursors.cursor(sc, hover=True)
                        @cursor.connect("add")
                        def on_add(sel, df=df, sheet_name=sheet_name, col=col):
                            i = sel.index
                            sel.annotation.set_text(
                                f"{sheet_name}\n{labels[i]}\n{df[time_col].iloc[i]}\n{df[col].iloc[i]:.2f}"
                            )

                plt.show()
            
            break

        # BWCS plotting
        elif station == "BWCS" or station == '2':
            BWCSdf = pd.DataFrame({
            "Timestamp": timestamp,
            "Cooler 1 IC Duties [1000-BTU/hr]": [q_BWC1_IC1, q_BWC1_IC2, q_BWC1_IC3, q_BWC1_IC4],
            "Cooler 2 IC Duties [1000-BTU/hr]": [q_BWC2_IC1, q_BWC2_IC2, q_BWC2_IC3, q_BWC2_IC4],
            "Cooler 3 IC Duties [1000-BTU/hr]": [q_BWC3_IC1, q_BWC3_IC2, q_BWC3_IC3, q_BWC3_IC4],
            "Cooler 4 IC Duties [1000-BTU/hr]": [q_BWC4_IC1, q_BWC4_IC2, q_BWC4_IC3, q_BWC4_IC4]
            })
            
            # writing data to excel sheet  
            sheetName = datetime.now().strftime("Run_%m%d%Y_%H%M%S") # labels sheet

            workbook = load_workbook("BWCSCoolers.xlsx")
            writer = pd.ExcelWriter("BWCSCoolers.xlsx", engine="openpyxl", mode="a")

            BWCSdf.to_excel(writer, sheet_name=sheetName, index=False)

            writer.close()

            print("\n\033[92m'BWCSCoolers.xlsx' now updated\033[0m\n")

            print("\n\033[92mNow generating plots...\033[0m")

            # generating plots for coolers
            # Load Excel sheets
            xls = pd.read_excel("BWCSCoolers.xlsx", sheet_name=None, engine="openpyxl")
            existing_df = next(iter(xls.values()))
            time_col = existing_df.columns[0]

            # Convert time column to datetime
            for name, df in xls.items():
                df[time_col] = pd.to_datetime(df[time_col])

            cooler_cols = [col for col in existing_df.columns if col != time_col]

            for col_index, col in enumerate(cooler_cols):
                    plt.figure(figsize=(8, 5))
                    scatter_objects = []

                    for sheet_name, df in xls.items():
                        x_vals = df[time_col]
                        y_vals = df[col]

                        # Scatter plot
                        sc = plt.scatter(x_vals, y_vals, label=sheet_name)
                        scatter_objects.append((sc, df, sheet_name, col))

                        labels = ['IC-1', 'IC-2', 'IC-3', 'IC-4']

                        # Optional: annotate points if needed
                        for i in range(len(x_vals)):
                            plt.annotate(
                                 labels[i],
                                 (x_vals.iloc[i], y_vals.iloc[i]),
                                 textcoords="offset points",
                                 xytext=(5, 5),
                                 ha='center',
                                 fontsize=8,
                                 color='gray'
                            )

                    # Define dashed lines once per figure
                    dashed_lines = {
                        'IC-1 Duty': 1100.4,
                        'IC-2 Duty': 1604.5,
                        'IC-3 Duty': 1759.5,
                        'IC-4 Duty': 2148.4,
                    }

                    # Draw dashed lines and label once
                    for label, y_val in dashed_lines.items():
                        plt.axhline(y=y_val, color='black', linestyle='--')
                        plt.text(
                            plt.gca().get_xlim()[0],
                            y_val + 12,
                            label,
                            color='black',
                            fontsize=7,
                            ha='left',
                            va='bottom'
                        )

                    plt.title(f"{col} vs {time_col}")
                    plt.xlabel(time_col)
                    plt.ylabel(col)
                    plt.grid(True)
                    plt.tight_layout()

                    # Add interactivity for this figure
                    for sc, df, sheet_name, col in scatter_objects:
                        cursor = mplcursors.cursor(sc, hover=True)
                        @cursor.connect("add")
                        def on_add(sel, df=df, sheet_name=sheet_name, col=col):
                            i = sel.index
                            sel.annotation.set_text(
                                f"{sheet_name}\n{labels[i]}\n{df[time_col].iloc[i]}\n{df[col].iloc[i]:.2f}"
                            )

            plt.show()

            break

        # bbcs plotting
        elif station == "BBCS" or station == '3':
            BBCSdf = pd.DataFrame({
            "Timestamp": timestamp,
            "Cooler 4 IC Duties [1000-BTU/hr]": [q_BBC4_IC1, q_BBC4_IC2, q_BBC4_IC3]
            })
            
            # writing data to excel sheet  
            sheetName = datetime.now().strftime("Run_%m%d%Y_%H%M%S") # labels sheet

            workbook = load_workbook("BBCSCoolers.xlsx")
            writer = pd.ExcelWriter("BBCSCoolers.xlsx", engine="openpyxl", mode="a")

            BBCSdf.to_excel(writer, sheet_name=sheetName, index=False)

            writer.close()

            print("\033[92m'BBCSCoolers.xlsx' now updated\033[0m\n")

            print("\n\033[92mNow generating plots...\033[0m")

            # generating plots for coolers
            # Load Excel sheets
            xls = pd.read_excel("BBCSCoolers.xlsx", sheet_name=None, engine="openpyxl")
            existing_df = next(iter(xls.values()))
            time_col = existing_df.columns[0]

            # Convert time column to datetime
            for name, df in xls.items():
                df[time_col] = pd.to_datetime(df[time_col])

            cooler_cols = [col for col in existing_df.columns if col != time_col]

            for col_index, col in enumerate(cooler_cols):
                    plt.figure(figsize=(8, 5))
                    scatter_objects = []

                    for sheet_name, df in xls.items():
                        x_vals = df[time_col]
                        y_vals = df[col]

                        # Scatter plot
                        sc = plt.scatter(x_vals, y_vals, label=sheet_name)
                        scatter_objects.append((sc, df, sheet_name, col))

                        labels = ['IC-1', 'IC-2', 'IC-3']

                        # Optional: annotate points if needed
                        for i in range(len(x_vals)):
                            plt.annotate(
                                 labels[i],
                                 (x_vals.iloc[i], y_vals.iloc[i]),
                                 textcoords="offset points",
                                 xytext=(5, 5),
                                 ha='center',
                                 fontsize=8,
                                 color='gray'
                            )

                    # Define dashed lines once per figure
                    dashed_lines = {
                        'IC-1 Duty': 1291,
                        'IC-2 Duty': 1762,
                        'IC-3 Duty': 1932,
                    }

                    # Draw dashed lines and label once
                    for label, y_val in dashed_lines.items():
                        plt.axhline(y=y_val, color='black', linestyle='--')
                        plt.text(
                            plt.gca().get_xlim()[0],
                            y_val + 12,
                            label,
                            color='black',
                            fontsize=7,
                            ha='left',
                            va='bottom'
                        )

                    plt.title(f"{col} vs {time_col}")
                    plt.xlabel(time_col)
                    plt.ylabel(col)
                    plt.grid(True)
                    plt.tight_layout()


                    # Add interactivity for this figure
                    for sc, df, sheet_name, col in scatter_objects:
                        cursor = mplcursors.cursor(sc, hover=True)
                        @cursor.connect("add")
                        def on_add(sel, df=df, sheet_name=sheet_name, col=col):
                            i = sel.index
                            sel.annotation.set_text(
                                f"{sheet_name}\n{labels[i]}\n{df[time_col].iloc[i]}\n{df[col].iloc[i]:.2f}"
                            )

            plt.show()

            break

    elif save == 'n':
        print("\n\033[92mExiting program...\033[0m\n")
        sys.exit()
    else:
        print("\nUnknown input. Please type 'Y' or 'N'")