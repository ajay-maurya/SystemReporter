# System Configuration Reporter

A comprehensive Windows system information tool that generates detailed HTML reports of hardware, software, and system configuration.

## Features

- **System Information**: OS details, hostname, IP address, MAC address
- **Hardware Details**: Manufacturer, model, serial number, motherboard information
- **Windows Information**: Edition, version, build, product ID, installation date
- **Security**: Masked Windows product key (only last 4 characters visible)
- **Microsoft Office Detection**: Version and installation path
- **CPU Information**: Cores, frequency, usage, model name
- **Memory Statistics**: RAM and swap usage details
- **Disk Information**: Partitions, usage statistics, I/O data
- **Network Configuration**: Interface details, addresses, I/O statistics
- **Process Monitoring**: Top 20 processes by CPU usage
- **Beautiful HTML Report**: Professional, styled output with timestamps

## Installation

### Option 1: Using the Pre-built Executable

1. Download the latest `SystemReporter.exe` from the releases page
2. Run the executable directly - no installation required

### Option 2: From Source Code

1. Ensure you have Python 3.6+ installed
2. Clone or download the source files
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Running the Application

**As an executable:**
```bash
SystemReporter.exe
```

**From source:**
```bash
python system_reporter.py
```

### Command Line Options

The application will:
1. Automatically collect system information
2. Generate an HTML report with timestamp in the filename
3. Optionally open the report in your default browser
4. Save the report in the current directory

### Output

The application creates an HTML file named `system_report_YYYYMMDD_HHMMSS.html` containing all system information in an easy-to-read format.

## Report Sections

1. **System Information**: Basic OS and network details
2. **Hardware Information**: Manufacturer, model, and serial numbers
3. **Windows Information**: Version, build, and licensing details
4. **Microsoft Office Information**: Installation status and version
5. **CPU Information**: Processor details and usage
6. **Memory Information**: RAM and swap statistics
7. **Disk Information**: Storage devices and usage
8. **Network Information**: Adapter configurations and statistics
9. **Running Processes**: Top processes by CPU usage

## Requirements

- Windows 7 or newer
- .NET Framework 4.5+ (usually included in modern Windows)
- For source version: Python 3.6+

## Building from Source

To create your own executable:

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   pip install pyinstaller
   ```

2. Build the executable:
   ```bash
   pyinstaller --onefile --console --name SystemReporter system_reporter.py
   ```

3. The executable will be in the `dist` folder

## Privacy and Security

- The application only reads system information - it does not modify any settings
- Sensitive information (like product keys) is masked in the report
- No data is transmitted over the network - all processing is local
- Reports are saved only to your local storage

## Troubleshooting

### Common Issues

1. **"Access denied" errors**: 
   - Run the application as Administrator for complete system information

2. **Missing hardware information**:
   - Some systems may not provide complete SMBIOS data

3. **Office not detected**:
   - The application detects Office via registry and file system
   - Older Office versions or non-standard installations might not be detected

### Getting Help

If you encounter issues:
1. Check that you're running a supported Windows version
2. Ensure you have necessary permissions to read system information
3. Verify that required dependencies are installed (for source version)

## License

This project is provided as-is for personal and educational use.

## Disclaimer

This tool is designed for system analysis and troubleshooting. Use responsibly and only on systems you own or have permission to inspect. The authors are not responsible for any misuse of this tool.
