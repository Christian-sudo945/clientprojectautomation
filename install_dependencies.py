import subprocess
import sys
import os
import urllib.request
import urllib.error
import zipfile
import io
import ssl
import time
import requests
import shutil

def install_package(package):
    """Install the specified package using pip."""
    print(f"Installing {package}...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
    except subprocess.CalledProcessError:
        print(f"Failed to install {package}. Trying with --user flag...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--user", package])

def download_edge_driver(version):
    """Download the Edge WebDriver and save it to edgedriver_win64 directory."""
    # Create driver directory if it doesn't exist
    driver_dir = os.path.join(os.path.dirname(__file__), 'edgedriver_win64')
    os.makedirs(driver_dir, exist_ok=True)
    
    # Extract major version
    major_version = version.split('.')[0]
    driver_version = None
    driver_url = None
    
    print(f"Finding WebDriver for Edge version {version} (major version: {major_version})...")
    
    # Try multiple methods to get the correct driver version
    
    # Method 1: Direct URL based on major version
    try:
        # Create a context with relaxed SSL verification 
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.110 Safari/537.36 Edg/96.0.1054.62',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        }
        
        # First, get the latest driver version for this Edge version
        driver_version_url = f"https://msedgedriver.azureedge.net/LATEST_RELEASE_{major_version}"
        print(f"Checking version URL: {driver_version_url}")
        
        req = urllib.request.Request(driver_version_url, headers=headers)
        
        with urllib.request.urlopen(req, timeout=15, context=ctx) as response:
            content = response.read()
            try:
                driver_version = content.decode('utf-8').strip()
                print(f"Found driver version: {driver_version}")
            except UnicodeDecodeError:
                print(f"Warning: Could not decode version response, trying fallback methods")
    except Exception as e:
        print(f"Method 1 failed: {e}")
    
    # Method 2: Try with requests library
    if not driver_version and 'requests' in sys.modules:
        try:
            print("Trying with requests library...")
            response = requests.get(
                f"https://msedgedriver.azureedge.net/LATEST_RELEASE_{major_version}",
                headers=headers,
                timeout=15,
                verify=False
            )
            if response.status_code == 200:
                driver_version = response.text.strip()
                print(f"Found driver version with requests: {driver_version}")
        except Exception as e:
            print(f"Method 2 failed: {e}")
    
    # Method 3: Use exact version for well-known Edge versions
    if not driver_version:
        # Map of some known Edge major versions to driver versions
        known_versions = {
            "135": "135.0.3179.98",
            "134": "134.0.3156.0",
            "133": "133.0.3151.18",
            "132": "132.0.3068.53",
            "131": "131.0.2998.0",
            "130": "130.0.2893.0",
            "129": "129.0.2781.0",
            # Add more mappings as needed
        }
        
        if major_version in known_versions:
            driver_version = known_versions[major_version]
            print(f"Using known driver version: {driver_version} for Edge {major_version}")
        else:
            # Fallback to using the exact Edge version
            driver_version = version
            print(f"Fallback: Using Edge version as driver version: {driver_version}")
    
    # Now try to download the driver
    driver_url = f"https://msedgedriver.azureedge.net/{driver_version}/edgedriver_win64.zip"
    temp_zip_path = os.path.join(driver_dir, "msedgedriver.zip")
    print(f"Driver download URL: {driver_url}")
    
    # Method 1: Download with urllib
    try:
        print(f"Downloading Edge WebDriver {driver_version}...")
        req = urllib.request.Request(driver_url, headers=headers)
        with urllib.request.urlopen(req, timeout=60, context=ctx) as response:
            with open(temp_zip_path, 'wb') as f:
                f.write(response.read())
        
        if os.path.exists(temp_zip_path) and os.path.getsize(temp_zip_path) > 0:
            print("Download successful with urllib.")
            # Extract the driver
            print("Extracting WebDriver...")
            with zipfile.ZipFile(temp_zip_path) as zip_ref:
                zip_ref.extractall(driver_dir)
            
            # Clean up
            os.remove(temp_zip_path)
            print(f"Edge WebDriver {driver_version} installed successfully!")
            return True
    except Exception as e:
        print(f"Download with urllib failed: {e}")
    
    # Method 2: Try with requests
    if 'requests' in sys.modules:
        try:
            print("Trying download with requests...")
            response = requests.get(driver_url, headers=headers, stream=True, verify=False, timeout=60)
            if response.status_code == 200:
                with open(temp_zip_path, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        f.write(chunk)
                
                if os.path.exists(temp_zip_path) and os.path.getsize(temp_zip_path) > 0:
                    print("Download successful with requests.")
                    # Extract the driver
                    print("Extracting WebDriver...")
                    with zipfile.ZipFile(temp_zip_path) as zip_ref:
                        zip_ref.extractall(driver_dir)
                    
                    # Clean up
                    os.remove(temp_zip_path)
                    print(f"Edge WebDriver {driver_version} installed successfully!")
                    return True
        except Exception as e:
            print(f"Download with requests failed: {e}")
    
    # Method 3: Try with PowerShell
    try:
        print("Attempting download with PowerShell...")
        command = [
            'powershell',
            '-Command',
            f"[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '{driver_url}' -OutFile '{temp_zip_path}' -UseBasicParsing"
        ]
        process = subprocess.run(command, capture_output=True, text=True)
        if process.returncode == 0 and os.path.exists(temp_zip_path) and os.path.getsize(temp_zip_path) > 0:
            print("Download successful with PowerShell.")
            # Extract the driver
            print("Extracting WebDriver...")
            with zipfile.ZipFile(temp_zip_path) as zip_ref:
                zip_ref.extractall(driver_dir)
            
            # Clean up
            os.remove(temp_zip_path)
            print(f"Edge WebDriver {driver_version} installed successfully!")
            return True
        else:
            print(f"PowerShell download failed: {process.stderr}")
    except Exception as e:
        print(f"Error using PowerShell for download: {e}")
    
    # Method 4: Direct download from Microsoft's WebDriver page
    try:
        print("Trying direct download from Microsoft WebDriver page...")
        # Check major version and use known direct download links
        direct_links = {
            "135": "https://msedgewebdriverstorage.blob.core.windows.net/edgewebdriver/135.0.3179.98/edgedriver_win64.zip",
            "134": "https://msedgewebdriverstorage.blob.core.windows.net/edgewebdriver/134.0.3156.0/edgedriver_win64.zip",
            # Add more as needed
        }
        
        if major_version in direct_links:
            direct_url = direct_links[major_version]
            print(f"Using direct Microsoft link: {direct_url}")
            
            response = requests.get(direct_url, headers=headers, stream=True, verify=False, timeout=60)
            if response.status_code == 200:
                with open(temp_zip_path, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        f.write(chunk)
                
                if os.path.exists(temp_zip_path) and os.path.getsize(temp_zip_path) > 0:
                    print("Download successful with direct link.")
                    # Extract the driver
                    print("Extracting WebDriver...")
                    with zipfile.ZipFile(temp_zip_path) as zip_ref:
                        zip_ref.extractall(driver_dir)
                    
                    # Clean up
                    os.remove(temp_zip_path)
                    print(f"Edge WebDriver {driver_version} installed successfully!")
                    return True
    except Exception as e:
        print(f"Direct download failed: {e}")
    
    print("All download methods failed.")
    return False

def get_edge_version():
    """Get the installed Microsoft Edge version."""
    # Method 1: Try registry (Windows)
    try:
        import winreg
        key_path = r"SOFTWARE\Microsoft\Edge\BLBeacon"
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path)
        value, _ = winreg.QueryValueEx(key, "version")
        winreg.CloseKey(key)
        print(f"Detected Edge version from registry: {value}")
        return value
    except Exception as e:
        print(f"Registry detection failed: {e}")

    # Method 2: Try executing Edge with --version flag
    try:
        edge_paths = [
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        ]
        
        for edge_path in edge_paths:
            if os.path.exists(edge_path):
                result = subprocess.run([edge_path, "--version"], 
                                       capture_output=True, text=True)
                if result.returncode == 0:
                    # Output format: "Microsoft Edge 135.0.3179.98"
                    version = result.stdout.strip().split(" ")[-1]
                    print(f"Detected Edge version from executable: {version}")
                    return version
    except Exception as e:
        print(f"Executable detection failed: {e}")
    
    print("Could not detect Microsoft Edge version through automatic methods.")
    return None

def main():
    """Install all required dependencies and WebDriver."""
    print("Starting setup process...")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    
    # Try to install requests first (for better download handling)
    try:
        install_package("requests")
        import requests
        # Disable SSL warnings
        requests.packages.urllib3.disable_warnings()
    except Exception as e:
        print(f"Could not install requests: {e}")
    
    print("\nInstalling required packages...")
    
    # Required Python packages
    required_packages = [
        "selenium",
        # tkinter can't be installed via pip as it comes with Python
    ]
    
    for package in required_packages:
        try:
            install_package(package)
        except Exception as e:
            print(f"Error installing {package}: {e}")
            if package == "selenium":
                print("Warning: Selenium is required for this application to work.")
    
    # Get and install Edge WebDriver
    edge_version = get_edge_version()
    if edge_version:
        print(f"\nDetected Microsoft Edge version: {edge_version}")
        if download_edge_driver(edge_version):
            print("\nAll dependencies installed successfully!")
        else:
            print("\nFailed to download Edge WebDriver automatically.")
            print("Please download it manually from:")
            print("https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/")
            
            # Give specific instructions
            driver_dir = os.path.join(os.path.dirname(__file__), 'edgedriver_win64')
            print(f"\nAfter downloading:")
            print(f"1. Extract the zip file")
            print(f"2. Copy msedgedriver.exe to: {driver_dir}")
            print(f"3. Ensure the file is named exactly 'msedgedriver.exe'")
    else:
        print("\nCould not detect Microsoft Edge.")
        print("Please make sure Microsoft Edge is installed on your system.")
        print("Then download the WebDriver manually from:")
        print("https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/")
    
    print("\nSetup process complete!")
    print("You can now run the main.py script to start the application.")
    time.sleep(5)  # Give user time to read message

if __name__ == "__main__":
    main()
