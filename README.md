
# Auto Clicker - Action Flow

## Overview
This Auto Clicker (Action Flow) is a versatile automation tool that allows users to automate repetitive tasks such as mouse clicks, key presses, and sequence of actions. It provides features like window targeting, customizable action intervals, action logging, and condition-based execution. This tool is built using Python with `customtkinter`, `keyboard`, `mouse`, `pywinauto`, and `pyautogui` libraries, and supports saving/loading presets for repetitive workflows.

## Features
<p align="center">  <img width="600" src="https://github.com/user-attachments/assets/34b1051d-01a9-4627-b46b-10f9c6d2b9a0" alt="ActionFlow"> </p>

- **Customizable Action Sequences:** Supports both mouse clicks and key presses, with optional delays between actions.
- **Target Window Actions:** Allows targeting specific windows for sending inputs, such as pressing keys or clicking the mouse.
- **Global Conditions:** Provides condition-based action execution, such as only running actions when a specific window is active.
- **Action Logging:** Logs actions to track what's being performed.
- **Preset System:** Supports saving and loading profiles for easy reuse of automation workflows.
- **Randomized Intervals:** Allows for setting random time intervals between actions.
- **Recording Mode:** Records user actions such as key presses and mouse clicks to build an automation sequence automatically.

## Requirements
The tool uses the following dependencies:
- `customtkinter`
- `keyboard`
- `mouse`
- `pyautogui`
- `pygetwindow`
- `pywinauto`
- `win32gui`
- `win32con`
- `win32api`
- `pynput`

Make sure to install these libraries before running the program.

### Installation
1. Clone the repository:
    ```bash
    git clone <repo_url>
    ```

2. Install dependencies using pip:
    ```bash
    pip install -r requirements.txt
    ```

### Running the Application
1. Run the main Python script:
    ```bash
    python ActionFlow.py
    ```

### Key Functionalities
1. **Creating Action Sequences:**
   - Add key press or mouse click actions with a set interval between them.
   - Optionally, randomize intervals between a range of values.

2. **Targeted Window Actions:**
   - Enable the option to send actions to a specific window, even if it is minimized or in the background.
   - Target a specific window from the active windows list for sending key presses or mouse clicks.

3. **Global Conditions:**
   - Set conditions that control when the actions will be performed, such as waiting for a specific window to become active.

4. **Recording Mode:**
   - Record key presses and mouse clicks in real-time, which are then saved into the action sequence for automation.

5. **Saving and Loading Presets:**
   - Save the current action sequence and other settings (such as the activation key and global condition) to a JSON file for reuse.
   - Load previously saved presets and restore all the settings for quick automation.

6. **Customization:**
   - Choose between different appearance modes (System, Light, Dark).
   - Adjust settings for window targeting and repeat options for action sequences.

### Controls
- **Set Activation Key:** Set a key to trigger starting and stopping the automation.
- **Start/Stop Automation:** Start or stop the automation sequence.
- **Set Global Condition:** Define conditions under which the automation will run.
- **Add Key Press/Mouse Click:** Add key press or mouse click actions to the sequence.
- **Record Actions:** Start recording key presses and mouse clicks to build an automation sequence.
- **Save/Load Profile:** Save the current sequence and settings as a preset, or load an existing preset.
- **Appearance Mode:** Change between light, dark, or system appearance.

### Configuration and Presets
The application will automatically load the last used preset when started if the configuration is available. Presets can be saved or loaded at any time.

#### Config File Structure (`config.json`):
- **last_preset**: Path to the last used preset file.

#### Preset File Structure (Example JSON):
```json
{
  "use_target_window": false,
  "activation_key": "f12",
  "action_sequence": [
    {
      "action_type": "key_press",
      "value": "a",
      "interval": 1.0,
      "min_interval": null,
      "max_interval": null,
      "target_window_title": null
    }
  ],
  "start_delay": 0.0,
  "repeat_option": "infinite",
  "repeat_count": null,
  "global_condition": null,
  "target_window_title": null
}
```

### How to Build with PyInstaller
To build a standalone executable with PyInstaller:
1. Install PyInstaller:
    ```bash
    pip install pyinstaller
    ```

2. Build the executable:
    ```bash
    pyinstaller --onefile --hidden-import=comtypes.stream ActionFlow.py
    ```

This will generate a standalone executable in the `dist` folder that can be run on any Windows machine.

## Contributing
Feel free to open issues or submit pull requests if you encounter bugs or want to contribute to the project.

## License
This project is licensed under the MIT License.
