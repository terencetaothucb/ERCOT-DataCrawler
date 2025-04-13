import os
import requests, time
from datetime import datetime
import pandas as pd

headers = {
    "accept": "*/*",
    "accept-language": "zh-CN,zh;q=0.9,en;q=0.8",
    "cache-control": "no-cache",
    "pragma": "no-cache",
    "priority": "u=1, i",
    "referer": "https://www.ercot.com/gridmktinfo/dashboards/ancillaryservices",
    "^sec-ch-ua": "^\\^Google",
    "sec-ch-ua-mobile": "?0",
    "^sec-ch-ua-platform": "^\\^Windows^^^",
    "sec-fetch-dest": "empty",
    "sec-fetch-mode": "cors",
    "sec-fetch-site": "same-origin",
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
    "x-requested-with": "XMLHttpRequest"
}

def crawler():
    result = []
    url = "https://www.ercot.com/api/1/services/read/dashboards/ancillary-services.json"
    response = requests.get(url, headers=headers)
    
    ###[DEBUG]###
    if response.text:
        print(response)
        try:
            data = response.json()
            # print(data)
            ascapmon = data['ascapmon']
        except ValueError:
            print("Response content is not valid JSON")
    else:
        print("Response is empty")
    ###[DEBUG]###
    
    ascapmon = response.json()['ascapmon']
    length = len(ascapmon)
    data = response.json()['data']
    for i in range(length):
        tagcLastTime = ascapmon[i].get('tagcLastTime')
        if tagcLastTime:
            tagcLastTime = datetime.fromtimestamp(tagcLastTime / 1000).strftime('%Y-%m-%d %H:%M:%S')
        deployedRegUp = ascapmon[i].get('deployedRegUp')
        deployedRegDown = ascapmon[i].get('deployedRegDown')
        undeployedRegUp = ascapmon[i].get('undeployedRegUp')
        undeployedRegDown = ascapmon[i].get('undeployedRegDown')
        rrs = ascapmon[i].get('rrs')
        nsrs = ascapmon[i].get('nsrs')
        ecrs = ascapmon[i].get('ecrs')
        currentFrequency = data[i].get('currentFrequency')

        # Collecting relevant information in a list
        info = [tagcLastTime, deployedRegUp, deployedRegDown, undeployedRegUp, undeployedRegDown, rrs, nsrs, ecrs, currentFrequency]
        print(info)
        result.append(info)

    # Get the date and time for file naming
    day = result[0][0].split(' ')[0].replace('-', '')
    st = result[0][0].split(' ')[-1].replace(':', '')
    et = result[-1][0].split(' ')[-1].replace(':', '')
    fileName = f'{day}-{st}-{et}'

    return result, fileName

def is_even_hour_and_zero_minute():
    """Check if the current time is an even hour exactly at zero minutes"""
    now = datetime.now()
    return now.hour % 2 == 0 and now.minute == 0

def is_past_even_hour():
    """Check if the current time is past an even hour by at least 3 minutes"""
    now = datetime.now()
    return now.minute >= 3

if __name__ == '__main__':
    # Wait until the top of the hour (even hour)
    while not is_even_hour_and_zero_minute():
        time.sleep(1)
    # Run the crawler task
    result, fileName = crawler()
    columns = ['Time', 'REG-UP-Deployed', 'REG-UP-Undeployed', 'REG-DOWN-Deployed', 'REG-DOWN-Undeployed', 'RRS', 'NON-SPIN', 'ECRS', 'Frequency']
    df = pd.DataFrame(result, columns=columns)
    
    # Define the folder path and file name
    folder_path = 'data/ercot'
    # Check if the folder exists, if not, create it
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    
    # Save the data to an Excel file
    df.to_excel(folder_path + f'/{fileName}-.xlsx', index=False)
    
    # Wait until 3 minutes after the even hour
    while not is_past_even_hour():
        time.sleep(1)