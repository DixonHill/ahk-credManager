# AHK CredManager

This script/class provides a graphical user interface (GUI) for securely managing credentials-such as usernames and passwords-using the Windows Credential Manager, all through AutoHotkey v2. It allows users to easily create, edit, delete, and view credentials via a convenient window that can be opened with a customizable hotkey. Credentials are stored with a specific prefix in Windows Credential Manager, ensuring they are kept secure and separate from other credentials. The script also exposes simple methods for other AutoHotkey scripts to programmatically retrieve or update credentials, making it a practical tool for both end users and developers who need to manage sensitive login information within their automation workflows.

Credit for the original class goes to [Droyo](https://www.autohotkey.com/boards/viewtopic.php?f=83&t=116285).

## Usage Instructions

1. **Download and Install AutoHotkey v2**  
   [Download AutoHotkey v2](https://www.autohotkey.com/) and install it on your system.

2. **Download the Script**  
   Clone this repository or download the script file to your computer.

3. **Run the Script**  
   Double-click the script file (`CredentialManager.ahk`) to launch the GUI.

4. **Manage Credentials**
   - **Add:** Click **Add** to create a new credential (username and password).
   - **Edit:** Select a credential and click **Edit** to modify it.
   - **Delete:** Select a credential and click **Delete** to remove it.
   - **View:** Select a credential to view its details.

5. **Access Credentials Programmatically**  
   Other AutoHotkey scripts can call the provided functions to retrieve or update credentials as needed.  
   You can include the class in other scripts using (`#include`).

7. **Hotkey**  
   Use the default hotkey (e.g., Shift+Ctrl+Alt+C) to open the credential manager window.  
   You can customize this hotkey in the script.

---

> **Note:**  
> All credentials are securely stored in Windows Credential Manager with a specific prefix to keep them separate from other credentials.

---

![image](https://github.com/user-attachments/assets/69176659-c78d-4506-877e-696d37bbb08f)
![image](https://github.com/user-attachments/assets/9bc521cd-e0d9-4616-bb78-001a9b38cf08)
![image](https://github.com/user-attachments/assets/88545755-d966-4332-b8d1-672fb39cd13d)
![image](https://github.com/user-attachments/assets/adc82a12-afd4-417b-958b-591eed052277)
