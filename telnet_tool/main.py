import telnetlib
import json
from datetime import datetime

def execute_telnet_command(ip, username, password, commands):
    try:
        tn = telnetlib.Telnet(ip)
        tn.read_until(b"Username: ")
        tn.write(username.encode('ascii') + b"\n")
        if password:
            tn.read_until(b"Password: ")
            tn.write(password.encode('ascii') + b"\n")
        for command in commands:
            tn.write(command.encode('ascii') + b"\n")
        output = tn.read_all().decode('ascii')
        tn.close()
        return output
    except Exception as e:
        return f"Error: {str(e)}"

def save_output_to_file(output, filename):
    with open(filename, 'w') as file:
        file.write(output)

def main():
    try:
        with open('telnet_setting.json') as config_file:
            config = json.load(config_file)
            dest_folder = config.get('dest', '')
            username = config.get('username', '')
            password = config.get('password', '')
            addresses = config.get('address', {})
            commands = config.get('cmd', [])
            
            for key, ip in addresses.items():
                output = execute_telnet_command(ip, username, password, commands)
                timestamp = datetime.now().strftime('%Y%m%d%H%M')
                filename = f"{dest_folder}/{key}-{timestamp}.txt"
                save_output_to_file(output, filename)
            print("Telnet commands executed successfully.")
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    main()
