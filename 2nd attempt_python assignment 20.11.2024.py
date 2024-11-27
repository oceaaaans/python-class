# -*- coding: utf-8 -*-
"""
Created on Wed Nov 20 17:02:18 2024

@author: ariel
"""

import subprocess
import sys
import csv
import socket
import platform
import os
import re
from datetime import datetime
import urllib.request


def check_install(package):
    try:
        __import__(package)
        print(f"{package} is available")
    except ImportError:
        print(f"{package} not available.\nInstalling now...")
        install(package)

packageList = ['uuid','psutil', 'pandas']

for package in packageList:
    check_install(package)


def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])
    
import psutil
import pandas
import uuid    
    
    
# Collects IP address using DNS resolution
def getIpAddress():
    try:
        hostName = socket.gethostname()
        ipAddress = socket.gethostbyname(hostName)
        return ipAddress
    except Exception as e:
        return f"Error: {e}"

# Collects MAC address using built-in system commands
def getMacAddress():
    try:
        if platform.system() == "Windows":
            result = subprocess.check_output("getmac", shell=True, text=True).splitlines()
            for line in result:
                if 'Physical' not in line and '-' in line:
                    mac_address = line.split()[0]
                    if len(mac_address) == 17:  # Valid MAC address length
                        return mac_address
        else:
            result = subprocess.check_output("ip link", shell=True, text=True)
            match = re.search(r"link/ether ([0-9a-fA-F:]{17})", result)
            if match:
                return match.group(1)
    except Exception as e:
        return f"Error: {e}"

    return "MAC address not found"

# Retrieves processor information using system tools more reliable across systems
def getProcessorInfo():
    try:
        if platform.system() == "Windows":
            command = "wmic cpu get name"
            result = subprocess.check_output(command, shell=True, text=True).splitlines()
            result = [line.strip() for line in result if line.strip()]
            return result[1] if len(result) > 1 else "N/A"
        elif platform.system() == "Linux":
            try:
                command = "lscpu | grep 'Model name' | awk -F: '{print $2}'"
                result = subprocess.check_output(command, shell=True, text=True).strip()
                if result:
                    return result.strip()
            except subprocess.CalledProcessError:
                command = "cat /proc/cpuinfo | grep 'model name' | uniq | awk -F: '{print $2}'"
                result = subprocess.check_output(command, shell=True, text=True).strip()
                return result.strip() if result else "N/A"
    except Exception as e:
        return f"Error: {e}"

    return "Processor model not found"

# Collects OS information
def getOsInfo():
    return f"{platform.system()} {platform.release()}"

# Collects system time in a readable format
def getSystemTime():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def getActivePorts():
    active_ports = []
    connections = psutil.net_connections(kind='tcp4')  # Fetch TCP connections
    for conn in connections:
        if conn.status == 'LISTEN':
            active_ports.append(conn.laddr.port)
    return active_ports

# Measures internet speed by checking connection response time
def getInternetSpeed():
    try:
        url = "http://google.com"
        startTime = datetime.now()
        urllib.request.urlopen(url, timeout=5)
        endTime = datetime.now()
        duration = (endTime - startTime).total_seconds()
        speedMbps = (1 / duration) * 8  # Approximate speed in Mbps for a 1MB request
        return f"{speedMbps:.2f} Mbps"
    except Exception:
        return "An error occurred"

# Collects computer information from the current system
def collectComputerInfo():
    computerName = socket.gethostname()
    ipAddress = getIpAddress()
    macAddress = getMacAddress()
    processorModel = getProcessorInfo()
    osInfo = getOsInfo()
    systemTime = getSystemTime()
    internetSpeed = getInternetSpeed()
    activePorts = getActivePorts()

    return {
        "Computer Name": computerName,
        "IP-address": ipAddress,
        "MAC-address": macAddress,
        "Processor Model": processorModel,
        "Operation System": osInfo,
        "System Time": systemTime,
        "Internet Connection Speed": internetSpeed,
        "Active Ports": activePorts}

# Writes or updates CSV with collected information
def writeToCsv(computerInfo, filename="computer_data_collection.csv"):
    fileExists = os.path.isfile(filename)
    updated = False

    if fileExists:
        with open(filename, mode='r', newline='') as file:
            reader = csv.DictReader(file)
            existingEntries = [row for row in reader]

    for entry in existingEntries if fileExists else []:
        if entry["Computer Name"] == computerInfo["Computer Name"] and entry["MAC-address"] == computerInfo["MAC-address"]:
            if entry != computerInfo:
                print("Warning: Existing entry differs from current file.")
                entry.update(computerInfo)
                updated = True
                break
            
         
            
    # Initialise the list for existing entries, in case the file doesn't exist
    existingEntries = []

    # If the file exists, read the existing entries
    if fileExists:
        with open(filename, mode='r', newline='') as file:
            reader = csv.DictReader(file)
            existingEntries = [row for row in reader]

    # Check if the computer is already in the CSV file
    computerExists = False
    for entry in existingEntries:
        if entry["Computer Name"] == computerInfo["Computer Name"] and entry["MAC-address"] == computerInfo["MAC-address"]:
            computerExists = True
            break

    # If the computer already exists, don't write it again
    if computerExists:
        print(f"Information for {computerInfo['Computer Name']} already exists. Skipping write.")
        return   



          
    # Combine active ports into a single string separated by semi-colons
    activePortsStr = "; ".join(map(str, computerInfo["Active Ports"]))

    # Open the file in 'a' (append) mode, check if file exists and update as needed
    with open(filename, mode='a', newline='') as file:
        writer = csv.writer(file)

        # If the file does not exist, write headers
        if not fileExists:
            writer.writerow(["Computer Name", "IP-address", "MAC-address", "Processor Model",
                             "Operation System", "System Time", "Internet Connection Speed", "Active Ports"])

        # Write the data (only one row with all the active ports)
        writer.writerow([
            computerInfo["Computer Name"],
            computerInfo["IP-address"],
            computerInfo["MAC-address"],
            computerInfo["Processor Model"],
            computerInfo["Operation System"],
            computerInfo["System Time"],
            computerInfo["Internet Connection Speed"],
            activePortsStr  # The active ports will be a single string separated by semicolons
        ])


# Main program
if __name__ == "__main__":
    info = collectComputerInfo()
    writeToCsv(info)
    print("Computer information collected and saved to CSV.")

