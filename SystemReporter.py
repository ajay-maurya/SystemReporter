#!/usr/bin/env python3
"""
Enhanced System Configuration Reporter
Generates an HTML report of system configuration including hardware details
"""

import platform
import socket
import psutil
import datetime
import os
import subprocess
import json
import uuid
import sys
import re
import winreg
from pathlib import Path

def get_system_info():
    """Get basic system information"""
    info = {
        'System': platform.system(),
        'Node Name': platform.node(),
        'Release': platform.release(),
        'Version': platform.version(),
        'Machine': platform.machine(),
        'Processor': platform.processor(),
        'Platform': platform.platform(),
        'Architecture': platform.architecture()[0],
        'Hostname': socket.gethostname(),
        'IP Address': socket.gethostbyname(socket.gethostname()),
        'MAC Address': ':'.join(['{:02x}'.format((uuid.getnode() >> elements) & 0xff) 
                               for elements in range(0, 8*6, 8)][::-1])
    }
    return info

def get_windows_info():
    """Get Windows-specific information"""
    try:
        # Get Windows version details
        win_ver = platform.win32_ver()
        
        # Try to get product key (masked)
        try:
            key = get_windows_product_key()
            masked_key = "XXXXX-XXXXX-XXXXX-XXXXX-" + key[-5:] if key else "Not available"
        except:
            masked_key = "Not available"
        
        # Get last update information
        last_update = get_last_update_time()
        
        windows_info = {
            'Windows Edition': win_ver[0],
            'Windows Version': win_ver[1],
            'Windows Build': win_ver[2],
            'Product ID': get_windows_product_id(),
            'Product Key': masked_key,
            'Last Update': last_update,
            'Installation Date': get_windows_install_date()
        }
    except Exception as e:
        windows_info = {'Error': f'Unable to retrieve Windows information: {str(e)}'}
    
    return windows_info

def get_windows_product_key():
    """Retrieve Windows product key (requires admin privileges)"""
    try:
        # Try to read from registry
        reg_path = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion"
        value_name = "DigitalProductId"
        
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path) as key:
            value, regtype = winreg.QueryValueEx(key, value_name)
            
        # Decode the product key from binary data
        key_map = "BCDFGHJKMPQRTVWXY2346789"
        key_chars = []
        product_id = list(value[52:67])
        
        for i in range(24, -1, -1):
            k = 0
            for j in range(14, -1, -1):
                k = (k << 8) | product_id[j]
                product_id[j] = k // 24
                k %= 24
            key_chars.append(key_map[k])
        
        # Format the key
        key = ''.join(key_chars[::-1])
        return '-'.join([key[i:i+5] for i in range(0, len(key), 5)])
    except:
        return None

def get_windows_product_id():
    """Get Windows Product ID"""
    try:
        reg_path = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion"
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path) as key:
            value, regtype = winreg.QueryValueEx(key, "ProductId")
            return value
    except:
        return "Not available"

def get_last_update_time():
    """Get last Windows update time"""
    try:
        # Check Windows Update history
        cmd = 'powershell "Get-WmiObject -Class Win32_QuickFixEngineering | Sort-Object InstalledOn -Descending | Select-Object -First 1"'
        result = subprocess.run(cmd, capture_output=True, text=True, shell=True)
        
        if result.returncode == 0 and result.stdout.strip():
            lines = result.stdout.split('\n')
            for line in lines:
                if 'InstalledOn' in line:
                    parts = line.split()
                    if len(parts) >= 3:
                        return ' '.join(parts[1:])
        
        return "Unknown"
    except:
        return "Unknown"

def get_windows_install_date():
    """Get Windows installation date"""
    try:
        cmd = 'systeminfo | find "Original Install Date"'
        result = subprocess.run(cmd, capture_output=True, text=True, shell=True)
        
        if result.returncode == 0:
            return result.stdout.split(':', 1)[1].strip()
        return "Unknown"
    except:
        return "Unknown"

def get_hardware_info():
    """Get hardware information (model, manufacturer, serial number)"""
    try:
        hardware_info = {}
        
        # Get computer system info
        try:
            cmd = 'wmic computersystem get model,manufacturer'
            result = subprocess.run(cmd, capture_output=True, text=True, shell=True)
            if result.returncode == 0:
                lines = result.stdout.strip().split('\n')
                if len(lines) > 1:
                    parts = lines[1].strip().split()
                    if len(parts) >= 2:
                        hardware_info['Manufacturer'] = parts[0]
                        hardware_info['Model'] = ' '.join(parts[1:])
        except:
            pass
        
        # Get BIOS serial number
        try:
            cmd = 'wmic bios get serialnumber'
            result = subprocess.run(cmd, capture_output=True, text=True, shell=True)
            if result.returncode == 0:
                lines = result.stdout.strip().split('\n')
                if len(lines) > 1:
                    hardware_info['Serial Number'] = lines[1].strip()
        except:
            pass
        
        # Get baseboard info
        try:
            cmd = 'wmic baseboard get product,manufacturer'
            result = subprocess.run(cmd, capture_output=True, text=True, shell=True)
            if result.returncode == 0:
                lines = result.stdout.strip().split('\n')
                if len(lines) > 1:
                    parts = lines[1].strip().split()
                    if len(parts) >= 2:
                        hardware_info['Motherboard Manufacturer'] = parts[0]
                        hardware_info['Motherboard Model'] = ' '.join(parts[1:])
        except:
            pass
        
        if not hardware_info:
            hardware_info = {'Error': 'Unable to retrieve hardware information'}
            
    except Exception as e:
        hardware_info = {'Error': f'Unable to retrieve hardware information: {str(e)}'}
    
    return hardware_info

def get_office_info():
    """Get Microsoft Office installation information"""
    try:
        office_info = {}
        
        # Check for Office in registry
        office_versions = {
            '16.0': 'Office 2016/2019/2021/365',
            '15.0': 'Office 2013',
            '14.0': 'Office 2010',
            '12.0': 'Office 2007',
            '11.0': 'Office 2003'
        }
        
        # Check common Office registry paths
        office_paths = [
            r"SOFTWARE\Microsoft\Office",
            r"SOFTWARE\Wow6432Node\Microsoft\Office"
        ]
        
        for office_path in office_paths:
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, office_path) as key:
                    for i in range(winreg.QueryInfoKey(key)[0]):
                        subkey_name = winreg.EnumKey(key, i)
                        if subkey_name in office_versions:
                            office_info['Version'] = office_versions[subkey_name]
                            
                            # Try to get more details
                            try:
                                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, f"{office_path}\\{subkey_name}\\Common\\InstallRoot") as install_key:
                                    path, _ = winreg.QueryValueEx(install_key, "Path")
                                    office_info['Install Path'] = path
                            except:
                                pass
                                
                            # Try to get product name
                            try:
                                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, f"{office_path}\\{subkey_name}\\Word\\InstallRoot") as word_key:
                                    path, _ = winreg.QueryValueEx(word_key, "Path")
                                    office_info['Detected Via'] = 'Word'
                            except:
                                try:
                                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, f"{office_path}\\{subkey_name}\\Excel\\InstallRoot") as excel_key:
                                        path, _ = winreg.QueryValueEx(excel_key, "Path")
                                        office_info['Detected Via'] = 'Excel'
                                except:
                                    office_info['Detected Via'] = 'Registry'
                            
                            break
            except:
                continue
        
        # If not found in registry, try to find in common install paths
        if not office_info:
            common_paths = [
                os.path.join(os.environ.get('ProgramFiles', 'C:\\Program Files'), 'Microsoft Office'),
                os.path.join(os.environ.get('ProgramFiles(x86)', 'C:\\Program Files (x86)'), 'Microsoft Office')
            ]
            
            for path in common_paths:
                if os.path.exists(path):
                    office_info['Version'] = 'Office (detected via file system)'
                    office_info['Install Path'] = path
                    break
        
        if not office_info:
            office_info['Status'] = 'Not detected'
            
    except Exception as e:
        office_info = {'Error': f'Unable to retrieve Office information: {str(e)}'}
    
    return office_info

def get_cpu_info():
    """Get CPU information"""
    try:
        cpu_info = {
            'Physical Cores': psutil.cpu_count(logical=False),
            'Total Cores': psutil.cpu_count(logical=True),
            'Max Frequency': f"{psutil.cpu_freq().max:.2f} MHz" if psutil.cpu_freq() else "N/A",
            'Current Frequency': f"{psutil.cpu_freq().current:.2f} MHz" if psutil.cpu_freq() else "N/A",
            'CPU Usage': f"{psutil.cpu_percent(interval=1)}%"
        }
        
        # Get CPU name
        try:
            cmd = 'wmic cpu get name'
            result = subprocess.run(cmd, capture_output=True, text=True, shell=True)
            if result.returncode == 0:
                lines = result.stdout.strip().split('\n')
                if len(lines) > 1:
                    cpu_info['Name'] = lines[1].strip()
        except:
            pass
            
    except:
        cpu_info = {'Error': 'Unable to retrieve CPU information'}
    
    return cpu_info

def get_memory_info():
    """Get memory information"""
    try:
        virtual_memory = psutil.virtual_memory()
        swap_memory = psutil.swap_memory()
        
        memory_info = {
            'Total RAM': f"{virtual_memory.total / (1024**3):.2f} GB",
            'Available RAM': f"{virtual_memory.available / (1024**3):.2f} GB",
            'Used RAM': f"{virtual_memory.used / (1024**3):.2f} GB",
            'RAM Usage': f"{virtual_memory.percent}%",
            'Total Swap': f"{swap_memory.total / (1024**3):.2f} GB",
            'Used Swap': f"{swap_memory.used / (1024**3):.2f} GB",
            'Swap Usage': f"{swap_memory.percent}%"
        }
    except:
        memory_info = {'Error': 'Unable to retrieve memory information'}
    
    return memory_info

def get_disk_info():
    """Get disk information"""
    try:
        disk_info = {}
        partitions = psutil.disk_partitions()
        
        for partition in partitions:
            try:
                usage = psutil.disk_usage(partition.mountpoint)
                disk_info[partition.device] = {
                    'File System': partition.fstype,
                    'Total Size': f"{usage.total / (1024**3):.2f} GB",
                    'Used': f"{usage.used / (1024**3):.2f} GB",
                    'Free': f"{usage.free / (1024**3):.2f} GB",
                    'Usage': f"{usage.percent}%",
                    'Mount Point': partition.mountpoint
                }
            except:
                continue
        
        # Add disk I/O statistics
        disk_io = psutil.disk_io_counters()
        if disk_io:
            disk_info['IO Statistics'] = {
                'Read Count': disk_io.read_count,
                'Write Count': disk_io.write_count,
                'Read Bytes': f"{disk_io.read_bytes / (1024**2):.2f} MB",
                'Write Bytes': f"{disk_io.write_bytes / (1024**2):.2f} MB"
            }
            
    except:
        disk_info = {'Error': 'Unable to retrieve disk information'}
    
    return disk_info

def get_network_info():
    """Get network information"""
    try:
        network_info = {}
        addresses = psutil.net_if_addrs()
        stats = psutil.net_if_stats()
        io_counters = psutil.net_io_counters(pernic=True)
        
        for interface, addrs in addresses.items():
            network_info[interface] = {
                'Addresses': [],
                'Is Up': stats[interface].isup if interface in stats else 'N/A',
                'Speed': f"{stats[interface].speed} Mbps" if interface in stats and stats[interface].speed else 'N/A'
            }
            
            for addr in addrs:
                network_info[interface]['Addresses'].append(
                    f"{addr.family.name}: {addr.address} (Netmask: {addr.netmask})"
                )
        
        # Add network I/O
        if io_counters:
            network_info['IO Statistics'] = {}
            for interface, io in io_counters.items():
                network_info['IO Statistics'][interface] = {
                    'Bytes Sent': f"{io.bytes_sent / (1024**2):.2f} MB",
                    'Bytes Received': f"{io.bytes_recv / (1024**2):.2f} MB",
                    'Packets Sent': io.packets_sent,
                    'Packets Received': io.packets_recv
                }
                
    except:
        network_info = {'Error': 'Unable to retrieve network information'}
    
    return network_info

def get_running_processes(limit=20):
    """Get list of running processes"""
    try:
        processes = []
        for proc in psutil.process_iter(['pid', 'name', 'username', 'memory_percent', 'cpu_percent']):
            try:
                processes.append(proc.info)
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
        
        # Sort by CPU usage and take top N
        processes.sort(key=lambda x: x['cpu_percent'] or 0, reverse=True)
        return processes[:limit]
        
    except:
        return [{'Error': 'Unable to retrieve process information'}]

def generate_html_report(data, output_path="system_report.html"):
    """Generate HTML report from collected data"""
    
    html_template = f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>System Configuration Report</title>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            line-height: 1.6;
            margin: 0;
            padding: 20px;
            background-color: #f5f5f5;
            color: #333;
        }}
        .container {{
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            padding: 30px;
            border-radius: 10px;
            box-shadow: 0 0 20px rgba(0,0,0,0.1);
        }}
        .header {{
            text-align: center;
            margin-bottom: 30px;
            padding-bottom: 20px;
            border-bottom: 3px solid #007acc;
        }}
        .header h1 {{
            color: #007acc;
            margin: 0;
        }}
        .timestamp {{
            color: #666;
            font-style: italic;
        }}
        .section {{
            margin-bottom: 30px;
            padding: 20px;
            border: 1px solid #ddd;
            border-radius: 8px;
            background: #fafafa;
        }}
        .section h2 {{
            color: #007acc;
            margin-top: 0;
            border-bottom: 2px solid #007acc;
            padding-bottom: 10px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 15px;
        }}
        th, td {{
            padding: 12px;
            text-align: left;
            border-bottom: 1px solid #ddd;
        }}
        th {{
            background-color: #007acc;
            color: white;
        }}
        tr:nth-child(even) {{
            background-color: #f2f2f2;
        }}
        tr:hover {{
            background-color: #e6f3ff;
        }}
        .warning {{
            color: #ff6b6b;
            font-weight: bold;
        }}
        .good {{
            color: #51cf66;
        }}
        .footer {{
            text-align: center;
            margin-top: 40px;
            padding-top: 20px;
            border-top: 2px solid #007acc;
            color: #666;
        }}
        .sensitive {{
            font-family: monospace;
            background-color: #ffe6e6;
            padding: 2px 5px;
            border-radius: 3px;
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🖥️ System Configuration Report</h1>
            <p class="timestamp">Generated on: {datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>
        </div>

        <!-- System Information -->
        <div class="section">
            <h2>📋 System Information</h2>
            <table>
                <tr>
                    <th>Property</th>
                    <th>Value</th>
                </tr>
                {"".join(f'<tr><td>{key}</td><td>{value}</td></tr>' for key, value in data['system_info'].items())}
            </table>
        </div>

        <!-- Hardware Information -->
        <div class="section">
            <h2>🔧 Hardware Information</h2>
            <table>
                <tr>
                    <th>Property</th>
                    <th>Value</th>
                </tr>
                {"".join(f'<tr><td>{key}</td><td>{value}</td></tr>' for key, value in data['hardware_info'].items())}
            </table>
        </div>

        <!-- Windows Information -->
        <div class="section">
            <h2>🪟 Windows Information</h2>
            <table>
                <tr>
                    <th>Property</th>
                    <th>Value</th>
                </tr>
                {"".join(f'<tr><td>{key}</td><td>{value if "Key" not in key else "<span class=\"sensitive\">" + value + "</span>"}</td></tr>' for key, value in data['windows_info'].items())}
            </table>
        </div>

        <!-- Office Information -->
        <div class="section">
            <h2>📊 Microsoft Office Information</h2>
            <table>
                <tr>
                    <th>Property</th>
                    <th>Value</th>
                </tr>
                {"".join(f'<tr><td>{key}</td><td>{value}</td></tr>' for key, value in data['office_info'].items())}
            </table>
        </div>

        <!-- CPU Information -->
        <div class="section">
            <h2>⚡ CPU Information</h2>
            <table>
                <tr>
                    <th>Property</th>
                    <th>Value</th>
                </tr>
                {"".join(f'<tr><td>{key}</td><td>{value}</td></tr>' for key, value in data['cpu_info'].items())}
            </table>
        </div>

        <!-- Memory Information -->
        <div class="section">
            <h2>💾 Memory Information</h2>
            <table>
                <tr>
                    <th>Property</th>
                    <th>Value</th>
                </tr>
                {"".join(f'<tr><td>{key}</td><td>{value}</td></tr>' for key, value in data['memory_info'].items())}
            </table>
        </div>

        <!-- Disk Information -->
        <div class="section">
            <h2>💿 Disk Information</h2>
            {"".join(f'''
            <h3>{disk}</h3>
            <table>
                <tr><th>Property</th><th>Value</th></tr>
                {"".join(f'<tr><td>{key}</td><td>{value}</td></tr>' for key, value in info.items())}
            </table>
            ''' for disk, info in data['disk_info'].items() if disk != 'IO Statistics')}
            
            {f'''
            <h3>Disk I/O Statistics</h3>
            <table>
                <tr><th>Property</th><th>Value</th></tr>
                {"".join(f'<tr><td>{key}</td><td>{value}</td></tr>' for key, value in data['disk_info'].get('IO Statistics', {}).items())}
            </table>
            ''' if 'IO Statistics' in data['disk_info'] else ''}
        </div>

        <!-- Network Information -->
        <div class="section">
            <h2>🌐 Network Information</h2>
            {"".join(f'''
            <h3>{interface}</h3>
            <table>
                <tr><th>Property</th><th>Value</th></tr>
                <tr><td>Status</td><td>{info.get("Is Up", "N/A")}</td></tr>
                <tr><td>Speed</td><td>{info.get("Speed", "N/A")}</td></tr>
                <tr><td>Addresses</td><td>{"<br>".join(info.get("Addresses", []))}</td></tr>
            </table>
            ''' for interface, info in data['network_info'].items() if interface != 'IO Statistics')}
            
            {f'''
            <h3>Network I/O Statistics</h3>
            <table>
                <tr><th>Interface</th><th>Bytes Sent</th><th>Bytes Received</th><th>Packets Sent</th><th>Packets Received</th></tr>
                {"".join(f'<tr><td>{interface}</td><td>{io["Bytes Sent"]}</td><td>{io["Bytes Received"]}</td><td>{io["Packets Sent"]}</td><td>{io["Packets Received"]}</td></tr>' 
                for interface, io in data['network_info'].get('IO Statistics', {}).items())}
            </table>
            ''' if 'IO Statistics' in data['network_info'] else ''}
        </div>

        <!-- Running Processes -->
        <div class="section">
            <h2>🔄 Top Running Processes (by CPU Usage)</h2>
            <table>
                <tr>
                    <th>PID</th>
                    <th>Name</th>
                    <th>User</th>
                    <th>CPU %</th>
                    <th>Memory %</th>
                </tr>
                {"".join(f'<tr><td>{proc.get("pid", "N/A")}</td><td>{proc.get("name", "N/A")}</td><td>{proc.get("username", "N/A")}</td><td>{proc.get("cpu_percent", "N/A")}</td><td>{proc.get("memory_percent", "N/A")}</td></tr>' 
                for proc in data['processes'])}
            </table>
        </div>

        <div class="footer">
            <p>Report generated by System Configuration Reporter</p>
            <p>© 2024 System Diagnostics Tool</p>
        </div>
    </div>
</body>
</html>
"""
    
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_template)
        return True
    except Exception as e:
        print(f"Error generating HTML report: {e}")
        return False

def main():
    """Main function"""
    print("🔍 System Configuration Reporter")
    print("=" * 40)
    print("Collecting system information...")
    
    # Collect all system information
    system_data = {
        'system_info': get_system_info(),
        'hardware_info': get_hardware_info(),
        'windows_info': get_windows_info(),
        'office_info': get_office_info(),
        'cpu_info': get_cpu_info(),
        'memory_info': get_memory_info(),
        'disk_info': get_disk_info(),
        'network_info': get_network_info(),
        'processes': get_running_processes()
    }
    
    # Generate output filename with timestamp
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"system_report_{timestamp}.html"
    
    print("Generating HTML report...")
    if generate_html_report(system_data, output_file):
        print(f"✅ Report generated successfully: {output_file}")
        print(f"📁 Location: {os.path.abspath(output_file)}")
        
        # Ask if user wants to open the report
        try:
            response = input("Would you like to open the report now? (y/n): ").lower()
            if response in ['y', 'yes']:
                os.startfile(output_file)  # This works on Windows
        except:
            print("Could not open the report automatically. Please open it manually.")
    else:
        print("❌ Failed to generate report")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nOperation cancelled by user")
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        input("\nPress Enter to exit...")
