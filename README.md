# VibrationVIEW Client

A Python-based client application for controlling VibrationVIEW software through file-based communication. This is a modern refactor of the original Visual Basic 6 application, providing the same functionality with improved reliability and maintainability.

## Overview

The VibrationVIEW Client enables remote control of VibrationVIEW software by sending commands through text files and monitoring response files. This allows automation and integration of vibration testing workflows.

## Features

- **Load Profiles**: Load VibrationVIEW random profiles (.vrp files)
- **Test Control**: Start and stop vibration tests remotely
- **Status Monitoring**: Query current system status
- **Data Conversion**: Convert VRD files to CSV format
- **Real-time Response**: Monitor command responses with automatic retry logic
- **Registry Integration**: Automatic detection of VibrationVIEW installation path

## Requirements

- Python 3.6 or higher
- Windows operating system
- VibrationVIEW software installed
- VibrationVIEW automation option (VR9604) - OR - VibrationVIEW may be run in Simulation mode without any additional hardware or software
- tkinter (included with Python)

## Installation

1. Ensure Python 3.6+ is installed on your system
2. Download `vibration_client.py`
3. No additional packages need to be installed (uses only standard library)

## Usage

### Running the Application

```
python vibration_client.py
```

### First Time Setup

1. The application will automatically detect your VibrationVIEW installation from the Windows registry
2. If the registry entry is not found, you'll be prompted to select the installation directory
3. The application will configure the remote control file path in the registry
4. **Important**: Restart VibrationVIEW after the first run to enable remote control

### Controls

- **Load**: Select and load a VibrationVIEW profile (.vrp file)
- **Run**: Start the currently loaded test
- **Stop**: Stop the running test
- **Status**: Query the current system status
- **Convert**: Convert VRD data files to CSV format (put in same folder as the data file)

### Response Monitoring

The application automatically monitors the response file (`RemoteControl.Status`) for changes and displays the results in the text area. It includes:

- Automatic retry logic (up to 3 retries)
- Timeout handling (3 seconds per command)
- Real-time status updates

## Configuration

### Version Configuration

The VibrationVIEW version is configured via a constant at the top of the file:

```python
VIBRATIONVIEW_VERSION = "2025.0"
```

Update this value to match your VibrationVIEW version if different.

### Registry Paths

The application reads from and writes to these registry locations:

- **Installation Path**: `HKEY_LOCAL_MACHINE\SOFTWARE\Vibration Research Corporation\{VERSION}`
- **Remote Control**: `HKEY_CURRENT_USER\SOFTWARE\Vibration Research Corporation\VibrationVIEW\{VERSION}\System Parameters`

### File Locations

- **Control File**: `RemoteControl.txt` (created in application directory)
- **Response File**: `RemoteControl.Status` (located in VibrationVIEW installation directory)

## Permissions

**Important Note**: The response file (`RemoteControl.Status`) is typically located in `C:\Program Files\...` which may require elevated permissions to access. You have two options:

1. **Run VibrationVIEW as Administrator** (Recommended)
2. **Run the client application as Administrator**
3. **Modify file permissions** on the VibrationVIEW installation directory

## Troubleshooting

### Common Issues

**"Registry Not Found"**
- VibrationVIEW may not be installed or the version constant needs updating
- Manually select the VibrationVIEW installation directory when prompted

**"No Host Response"**
- Ensure VibrationVIEW is running
- Verify VibrationVIEW was restarted after first client run
- Check that VibrationVIEW has permissions to write to the response file
- Try running VibrationVIEW as Administrator

**"Response file not found"**
- VibrationVIEW may not be configured for remote control
- Check the installation path is correct
- Check that VibrationVIEW has required elevated permissions to access this file

**Permission Errors**
- Run VibrationVIEW as Administrator
- Alternatively, run the client application as Administrator

### Debug Information

The application provides console output for debugging:
- Registry read/write operations
- File path configuration
- Command sending status

## Technical Details

### Communication Protocol

The application uses a simple file-based protocol:

1. **Commands** are written to `RemoteControl.txt`
2. **Responses** are read from `RemoteControl.Status`
3. **File modification time** is monitored to detect new responses

### Supported Commands

- `load <filepath>` - Load a profile file
- `run` - Start the test
- `stop` - Stop the test
- `status` - Get system status

### Error Handling

- Automatic retry logic for failed commands
- Graceful handling of file access errors
- User-friendly error messages
- Registry operation error recovery

