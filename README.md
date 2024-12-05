# Wi-Fi Password Extractor (VBScript)

This VBScript retrieves the password for a specified Wi-Fi network on a Windows system and saves it to a Notepad file (`passwort.txt`). It uses the `netsh wlan show profiles` command to extract the password in plaintext.

## ⚠️ Disclaimer
This script is for **educational purposes only**. Ensure you have proper authorization to retrieve Wi-Fi passwords before using this script. Unauthorized use may violate privacy and legal regulations.

---

## Features
- Extracts the password for a specific Wi-Fi network.
- Automates the process using VBScript and Command Prompt.
- Saves the password into a text file (`passwort.txt`) for easy reference.

---

## Prerequisites
1. **Windows Operating System**: This script is designed for Windows environments.
2. **Administrative Privileges**: Running the script requires admin rights.
3. **Existing Wi-Fi Profile**: The Wi-Fi network must be saved on the system.

---

## Usage Instructions
1. Clone or download this repository.
2. Open the script file and edit the following line to set the target Wi-Fi network:
   ```vbscript
   wifiName = "your-wifi-network-name"
   ```
3. Save the file as `WiFiPasswordExtractor.vbs`.
4. Double-click the file to run it. If prompted, grant administrative privileges.
5. The script will:
   - Open Command Prompt and retrieve the password.
   - Save the password in a file named `password.txt` in the script's directory.

---

## Troubleshooting
If the script doesn't work as expected:
- Ensure the target Wi-Fi network name (`wifiName`) matches the profile stored on your system.
- Run the script with administrative privileges.
- Disable any antivirus or security tools that may block the script.

---

## Limitations
- The script relies on `SendKeys`, which can be affected by other windows stealing focus.
- The password is temporarily stored in the clipboard, which may overwrite existing clipboard content.

---

## Legal Notice
This script is intended for retrieving passwords of Wi-Fi networks you own or have explicit permission to access. Unauthorized access to networks may violate local laws.

---

## License
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
```
