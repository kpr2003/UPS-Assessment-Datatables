import requests, json, time
from dotenv import load_dotenv
import os
import pandas as pd

load_dotenv()

# Configuration variables
API_URL = os.getenv('API_URL')
API_KEY = os.getenv('API_KEY')
ORG_ID = os.getenv('ORG_ID')
OFFSET = 0
HEADERS = {
        'Accept': 'application/json',
        'Authorization': f'Bearer {API_KEY}'
    }

def get_org_id ():
    try:
        endpoint = f'{API_URL}/organizations'
        response = requests.get(endpoint, headers=HEADERS)
        response.raise_for_status()
        org_id = json.loads(response.text)["organizationIds"][0]
        return org_id
    except requests.exceptions.RequestException as e:
        print(f"Error fetching the organization id: {e}")
        return None
    
def get_org_inv (org_id):
    try:
        endpoint = f'{API_URL}/organizations/{org_id}/inventory'
        response = requests.get(endpoint,headers=HEADERS)
        response.raise_for_status()
        inv_data = response.json()
        with open (f'org_inv.json', 'w') as json_file:
            json.dump(inv_data, json_file, indent=4)
        print('Inventory data written to org_inv.json')
    except requests.exceptions.RequestException as e:
        print(f"Error fetching the organization inventory: {e}")
    
def get_sensors(org_id):
    try:
        endpoint = f'{API_URL}/organizations/{org_id}/sensors'
        response = requests.get(endpoint, headers=HEADERS)
        response.raise_for_status()
        sensor_data = response.json()
        with open(f'sensor_inv.json', 'w') as json_file:
            json.dump(sensor_data, json_file, indent=4)
        print('Sensor data written to sensor_inv.json')
    except requests.exceptions.RequestException as e:
        print(f"Error fetching the sensor data: {e}")

def get_measurements(org_id):
    try:
        endpoint = f'{API_URL}/organizations/{org_id}/measurements/live'
        start = time.time()
        with requests.get(endpoint, headers=HEADERS, stream=True) as response:
            response.raise_for_status()
            # Open a file to write the streaming data
            with open(f'measurements.json', 'w') as json_file:
                json_file.write('[')
                first_line = True
                for line in response.iter_lines():
                    # Filter out keep-alive new lines
                    if line:
                        # print(line)
                        # Parse the JSON document
                        json_data = json.loads(line)
                        # Write the JSON data to the file
                        if not first_line:
                            json_file.write(',\n')
                        json.dump(json_data, json_file, indent=4)
                        first_line = False
                    if time.time() - start > 400:
                        break
                json_file.write(']')
            print(f'Live measurements data written to measurements.json')
    except requests.exceptions.RequestException as e:
        print(f"Error fetching live measurements: {e}")

def filter_org_inv():
    with open('org_inv.json') as file:
        data = json.load(file)

    inv = data['inventoryObjects']

    locations = {}
    for location in inv:
        if 'inventoryObjectType' in location and location['inventoryObjectType'] == 'Location' and 'address' in location:
            locations[location['id']] = [location['label'], location['address']]
        elif 'inventoryObjectType' in location and location['inventoryObjectType'] == 'Location':
            locations[location['id']] = [location['label']]

    ups_inv = {}
    for device in inv:
        if 'type' in device and device['type'] == 'UPS':
            if device['locationId'] in locations:
                ups_inv[device['id']] = [device['label'], 
                                         locations[device['locationId']], 
                                         device['modelName'], 
                                         device['ipV4Addresses'][0],
                                         device['firmwareVersion'],
                                         device['serialNumber']]
                # print(ups_inv[device['id']])
     
        # if 'type' in device and device['type'] == 'UPS':
        #     id = device['id']
        #     label = device['label']
        #     #type = device['type']
        #     #manufacturer = device['manufacturerName']
        #     #serial_number = device['serialNumber']
        #     #firmware_version = device['firmwareVersion']
        #     #ip_addresses = device['ipV4Addresses']
        #     hostname = device['hostname']

        #     'id': id,
        #     'label': label,
        #     #'type': type
        #     #'manufacturer': manufacturer,
        #     #'serial_number': serial_number,
        #     #'firmware_version': firmware_version,
        #     #'ip_addresses': ip_addresses,
        #     'hostname': hostname

    with open('sensor_inv.json') as file:
        sensors_inv = json.load(file)

    sensors = sensors_inv['sensors']

    sensor_dict = {}
    for sensor in sensors:
        sensor_dict[sensor['id']] = [sensor['deviceId'], sensor['name'], sensor['unit']]

    with open('measurements.json') as file:
        live_data = json.load(file)

    for data in live_data:
        if data['sensorId'] in sensor_dict.keys():
            sensor_dict[data['sensorId']].append(data)

    for key, value in sensor_dict.items():
        if value[0] in ups_inv.keys():
            ups_inv[value[0]].append(sensor_dict[key][1:])
    
    # Take the first two elements (location information) and concatenate with the remaining elements which are sorted
    sorted_ups_inv = {key: value[:6] + sorted(value[6:], key=lambda x: x[0]) for key, value in ups_inv.items()}

    # Writing data to a JSON file
    with open('test_writing.json', 'w') as f:
        json.dump(sorted_ups_inv, f, indent=4)


def tt():

    # Load JSON data
    with open('test_writing.json') as f:
        data = json.load(f)

    # Initialize an empty list to store rows
    rows = []

    # Extract sensor names from the first site entry to set the columns
    first_site = next(iter(data.values()))
    sensor_names = [sensor[0] for sensor in first_site[6:]]

    # Iterate through each site in the JSON data
    for device_id, device_info in data.items():
        device_name = device_info[0]    # Device Name
        site_name = device_info[1][0]   # Site name
        device_model = device_info[2]   # Model Name
        device_ip_addr = device_info[3] # IP Address
        device_firm_v = device_info[4]  # Firmware Version
        device_ser_num = device_info[5] # Serial Number
        # print([device_name,site_name,device_model,device_ip_addr, device_firm_v, device_ser_num])

        row = {'Device Name': device_name, 'Site Name': site_name, 
               'Model': device_model, 'IP Address': device_ip_addr,
               'Firmware Version': device_firm_v, 'Serial Number': device_ser_num
               }
        for sensor in device_info[6:]:
            print(sensor)
            sensor_name = sensor[0]
            if len(sensor) > 2 and 'numericValue' in sensor[2]:
                row[sensor_name] = sensor[2]['numericValue']
            elif len(sensor) > 2 and 'stringValue' in sensor[2]:
                row[sensor_name] = sensor[2]['stringValue']
            else:
                row[sensor_name] = None
        rows.append(row)

    # Create a DataFrame from the rows
    df = pd.DataFrame(rows, columns=['Device Name', 'Site Name','Model',
                                     'IP Address','Firmware Version',
                                     'Serial Number'] + sensor_names)

    # Write the DataFrame to an Excel file
    output_path = 'test.xlsx'

    if os.path.exists(output_path):
        os.remove(output_path)
    df.to_excel(output_path, index=False)
    print(f"Data has been written to {output_path}")


if __name__ == '__main__':

    try:
        # Fetches the organization id
        # org_id = get_org_id()
        # print(org_id)

        # prints the constant (organization id)
        # print(ORG_ID)
        
        # Pulls the organization inventory
        # get_org_inv(ORG_ID)
        
        # Pulls the sensor ids of devices
        # get_sensors(ORG_ID)
        
        # Fetches the live measurements of all sensors randomly
        # get_measurements(ORG_ID)

        # filter_org_inv()
        tt()
        
        # Convert JSON to CSV
        #json_to_csv('test_writing.json', 'ups_data.csv')

    except KeyboardInterrupt:
        print("\nProcess interrupted.")
