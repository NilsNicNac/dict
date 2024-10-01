import pandas as pd

def dictToExcel(dicts, filename):
    # Group data by Session_Id
    grouped_data = {}
    for dict in dicts:
        session_id = dict.get("Session_Id")
        if session_id not in grouped_data:
            grouped_data[session_id] = {
                "createdAt": [],
                "_id": [],
                "latitude": [],
                "longitude": [],
                "temperature": [],
                "co2_level": [],
                "dust_particles": [],
                "ctr": [],
                "updatedAt": [],
            }
        location = dict.get("location", {})
        grouped_data[session_id]["createdAt"].append(dict.get("createdAt"))
        grouped_data[session_id]["_id"].append(dict.get("_id"))
        grouped_data[session_id]["latitude"].append(location.get("latitude"))
        grouped_data[session_id]["longitude"].append(location.get("longitude"))
        grouped_data[session_id]["temperature"].append(dict.get("temperature"))
        grouped_data[session_id]["co2_level"].append(dict.get("co2_level"))
        grouped_data[session_id]["dust_particles"].append(dict.get("dust_particles"))
        grouped_data[session_id]["ctr"].append(dict.get("ctr"))
        grouped_data[session_id]["updatedAt"].append(dict.get("updatedAt"))
    
    # Create a Pandas Excel writer using XlsxWriter as the engine
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        for session_id, output in grouped_data.items():
            df = pd.DataFrame(output)
            sheet_name = f"Session_Id {session_id}"
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"Data successfully saved to {filename}")

# test input
dictToExcel([{
  "_id": "66f5a22f271ce1b69b2460f4",
  "location": {
    "latitude": 51.02989578,
    "longitude": 4.475098133
  },
  "temperature": 22.5,
  "co2_level": 412,
  "dust_particles": 12,
  "ctr": 0,
  "createdAt": "2024-09-26T18:04:31.608Z",
  "updatedAt": "2024-09-26T18:04:31.608Z",
  "Session_Id": 1
},{
  "_id": "66f5a22f271ce1b69b2460f4",
  "location": {
    "latitude": 52.02989578,
    "longitude": 4.475098133
  },
  "temperature": 21.5,
  "co2_level": 312,
  "dust_particles": 22,
  "ctr": 0,
  "createdAt": "2024-09-22T18:04:31.608Z",
  "updatedAt": "2024-09-22T18:04:31.608Z",
  "Session_Id": 2
}], "output.xlsx")