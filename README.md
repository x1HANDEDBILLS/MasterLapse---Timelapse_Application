# **MasterLapse | Timelapse PowerShell Application**


![image](https://github.com/user-attachments/assets/e96528af-d0fe-4fed-8b47-41934bd63daa)



## **Overview**
This Timelapse PowerShell Script is an automated solution for capturing images from a PC and webcam at set intervals and turning them into a timelapse video. It runs in the background, manages files automatically, and even restarts itself if necessary. For Windows 10/11

---

## **What This Script Can Do**

### **Automated Image Capture and Video Creation**
- Takes pictures at set intervals (seconds or minutes).
- Works with most webcams at high resolutions.
- **Captures PNG lossless quality images for the best output.**
- Converts captured images into MP4 video using FFmpeg after each recording session.

### **Dark Frame Removal**
- Automatically removes dark/black frames for better quality.
- Adjustable sensitivity setting from 0 to 255:
  - **0:** No images are removed.
  - **1:** Removes the darkest images.
  - **20:** A good starting point for most users.
  - Increasing the number makes it more sensitive, removing progressively lighter dark images.
- Useful for plant timelapses, removing night-time images for a cleaner final video.
- Helps eliminate accidental black frames that can appear during a capture session.

### **Auto Restart Functionality (Optional)**
- If enabled, the script will automatically restart if it crashes or is accidentally closed, and it will immediately resume recording.
- If your computer reboots, the script will start itself and begin recording automatically with no user input required.
- Ensures continuous operation without manual intervention.
- If enabled, merges multiple videos into a master video once a set limit is reached.
- Can be set to automatically restart after a specified amount of time (minutes, hours, or days) to ensure uninterrupted operation.

### **Graphical User Interface (GUI)**
- Live camera preview with a fullscreen preview option.
- **All necessary controls including:**
  - Start/Stop recording.
  - Set camera resolution.
  - Configure capture intervals.
  - Adjust final video framerate.
  - Additional customizable settings.

### **Timelapse Calculator**
- Estimates the total number of frames captured based on the selected capture interval and capture duration.
- Displays the estimated final video duration based on the chosen framerate.
- **Displays the estimated final video size.**
- Helps users plan their time-lapse sessions efficiently.

#### **Example Calculation:**
| **Parameter**             | **Value**                                                  |
| --------------------- | ------------------------------------------------------ |
| Capture Interval      | 10 seconds                                             |
| Capture Duration      | 2 hours (7200 seconds)                                 |
| Total Frames Captured | 720 frames                                             |
| Final Video Framerate | 30 FPS                                                 |
| Final Video Duration  | 24 seconds                                             |
| Estimated Video Size  | \~XX MB (based on resolution and compression settings) |

**Explanation:**
- The script captures an image every 10 seconds.
- Over a 2-hour duration, there are 7200 seconds.
- The total frames captured would be **7200 seconds / 10 seconds = 720 frames**.
- If the final video is set to **30 frames per second (FPS)**, the resulting time-lapse video will be **720 frames / 30 FPS = 24 seconds long**.
- The estimated file size depends on resolution, quality, and compression settings.

---

## **Logging and Error Handling**
- Keeps a log of all actions for troubleshooting and review.

---

## **Requirements**

### **Software**
- **FFmpeg:** Must be installed and added to the system PATH.
- **AForge.NET Framework:** Required for webcam functionality.
- **PowerShell:** Version 5.1 and 7 or later.

### **Hardware**
- **Webcam:** Any compatible webcam.

---

## ⚠️ **WARNING:**
- The window may appear **unresponsive** during video creation and video merging. This is a normal **temporary freeze** until the process completes.
- **Avoid capturing more than 2000 images per video** to prevent excessive memory usage and potential script crashes.
- **Keep the number of merged videos low** to avoid script instability.
- **Auto Restart and Auto Merge Video options can help manage large files efficiently.**
- Currently working on improving stability for handling larger image and video sets.

---

## **How to Install**

### **Automatic Installation Method**

1. Click the green **"Code"** button on the repository page and download the .zip file.
2. Your browser may flag the file as "Not Trusted." If this happens:
   - Click **Keep** when prompted.
   - Then, select **Keep Anyway** to allow the download.
3. **Extract the Files:**
   - Locate the downloaded .zip file and extract it to your desired location.
4. **Run the Installer:**
   - Find **Masterlapse Installer.exe** inside the extracted folder.
   - Double-click to launch it.
   - Follow all prompts—no changes are necessary during installation.
   - Open the Application via Desktop Shortcut.

### **What the Installer Does**
- Ensures it is running in Windows PowerShell 5.1.
- Requests Administrator privileges if needed.
- Installs FFmpeg if not already installed.
- Installs PowerShell 7 if not already installed.
- Installs AForge.NET Framework (2.2.5) if not already installed.
- Creates and configures scheduled tasks to ensure auto-restart functionality.
- Adds a Desktop Shortcut for quick access.

### **Manual Installation Method**

1. **Install PowerShell 7**
   ```powershell
   winget install Microsoft.Powershell --silent
   ```

2. **Install FFmpeg**
   ```powershell
   winget install ffmpeg
   ```

3. **Install AForge.NET Framework**
   - Download AForge.NET Framework from the official AForge.NET Downloads page.
   - Install version **2.2.5** using the provided installer.
   - **Note:** Use the default installation path.

4. **Ensure Folder Placement**
   - Place the Timelapse folder (which contains **Images, Videos, Logs Folders, and a .ICO**) in:
     ```
     C:\Users\User
     ```
   - Replace "User" with your actual Windows username and delete the DELETE.txt files inside each **Images, Videos, and Logs Folder**.

5. **Ensure Folder Placement**
   - Copy the Timelapse folder from Step 4 and paste it in:
     ```
     C:\Program Files (x86)
     ```
   - Delete the **Images, Videos, and Logs folders**. This folder should only contain a **.ico** image.

6. **Ensure Script Placement**
   - Place `TimeLapse-v1.0.0.ps1` in:
     ```
     C:\Program Files (x86)\Timelapse
     ```

7. **Start the Script Manually**
   ```powershell
   cd "C:\Program Files (x86)\Timelapse"
   .\TimeLapse-v1.0.0.ps1
   ```

---

## **Troubleshooting**

### **Common Issues**

- **No Webcam Detected**: Ensure your webcam is connected and recognized by the system. Verify AForge.NET is properly installed.
- **FFmpeg Not Found**: Confirm FFmpeg is installed and added to the system PATH.
- **Permissions Error**: Run PowerShell as Administrator to ensure proper access.
- **Camera Not Switching**: Ensure the selected device is not in use by another application.

---

## **License**
This project is licensed under the MIT License.

---

## **Acknowledgements**
- **AForge.NET Framework** for camera handling.
- **FFmpeg** for video encoding.

**Author:**
**x1HANDEDBILLS** - For questions or support, please open an issue or contact me via GitHub.

