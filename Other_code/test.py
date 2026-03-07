import requests
 
statusURL = "http://172.30.73.149:1810/ScaMonitor/GetMachineStatus_All?productionDate=2026-01-26"

countURL = "http://172.30.73.149:1810/ScaMonitor/GetInspectionOkNg_All?productionDate=2026-01-26"
 
def get_machine_status_pie_all(production_date: str):

    params = {"productionDate": production_date}

    response = requests.get(statusURL, params=params, timeout=10)

    response.raise_for_status()  

    return response.json()
 
def get_inspection_ok_ng_all(production_date: str):

    params = {"productionDate": production_date}

    resp = requests.get(countURL, params=params, timeout=10)

    resp.raise_for_status()

    return resp.json()
 
if __name__ == "__main__":

    #Dữ liệu status

    data1 = get_machine_status_pie_all("2026-01-20")

    print(data1)

    data2 = get_inspection_ok_ng_all("2026-01-20")
    print(data2)
 