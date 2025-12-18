# RemotePCInfo - Remote Computer Information Gathering Script

Ссылка на русскую документацию - 

#### RemotePCInfo is a PowerShell script for automatically collecting system information from remote computers on a network. The script uses PsExec for remote command execution and gathers detailed data about hardware and software configuration.

Features
- 📊 System and hardware information collection
- 🌐 Remote access support via PsExec
- 📁 Automatic result saving to file
- ✅ Computer availability check before scanning
- 📋 Batch processing from a file with IP addresses list
- 📈 Detailed execution statistics

---

- Collected Information
    - General System Information
    - Computer Name
    - Domain

- Operating System Age Information
    - Operating System
    - Version
    - Installation Date
    - Last Boot Time
    - Current Local Time
    - Days Since Installation

- Motherboard
    - Manufacturer
    - Model
    - Serial Number

- Processor
    - Model
    - Number of Cores
    - Maximum Clock Speed

- RAM
    - Information about each module
    - Manufacturer
    - Size of each module
    - Total capacity

- Disks
    - Model
    - Size
    - Interface Type
    - Media Type
    - Serial Number

- Video Adapter
    - Video card name
    - Video memory size
    - Driver version
    - Video processor

---

### Requirements
- Windows PowerShell 5.1 or higher
- Administrative rights on remote computers
- Network access to remote computers
- PsExec from Sysinternals Suite

---

### Installation
1. Clone the repository or download the script:
```
git clone https://github.com/ваш-репозиторий/RemotePCInfo.git
```
2. Place one of the PsExec files in the script folder:

- PsExec64.exe (рекомендуется)
- PsExec.exe

3. The computers.txt file will be created automatically on first run

---

### Project Structure
### Структура проекта
```
RemotePCInfo/
├── RemotePCInfo.ps1          # Main script
├── README.md                 # Russian documentation
├── README_ENG.md             # English documentation
├── PsExec64.exe              # PsExec 64-bit
├── PsExec.exe                # PsExec 32-bit
├── computers.txt             # Computers list (created automatically)
└── CollectedPCInfo.txt       # Scanning results (created automatically)
```

---

### Usage
#### Preparation

1. Add PsExec to the script folder
2. Edit the computers.txt file, adding target computer IP addresses

#### Example computers.txt content:
```
# File with list of computers for scanning
# Each IP address on a new line
# Lines starting with # are ignored

192.168.1.100
192.168.1.101
# 192.168.1.102 - this computer is commented out
```
### Launch
1. Launch PowerShell console as Administrator.
2. Navigate to the directory where the project is located.
```
cd <your-directory-path>
```
3. Run the script
```
.\RemotePCInfo.ps1
```

### Configuration
- Connection Parameters
The script uses PsExec with the following parameters:
```
-accepteula - automatic acceptance of license agreement
-nobanner - hides PsExec banner
```

### Network Requirements
- Must have open ports
```
445 (SMB) 
135 (RPC)
```
- Account must have administrative rights on remote computers
- Domain/workgroup must be properly configured

### Limitations
- Requires administrative rights
- Works only in Windows domains or workgroups
- Does not support Linux/macOS
- Requires PowerShell installed on remote computers

### Security
- The script uses standard Windows tools
- Does not transmit data to the internet
- All data is saved locally
- Recommended for use in trusted networks

### Troubleshooting
#### Computer unavailable (ping)
- Check network connection
- Ensure the computer is powered on
- Check firewall settings

#### PsExec Errors
- Ensure PsExec is in the script folder
- Check administrative rights
- Ensure remote computer is network accessible

#### WMI Access Errors
- Check Windows Management Instrumentation service
- Verify administrative rights