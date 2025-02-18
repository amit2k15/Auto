from pyzabbix import ZabbixAPI
import pandas as pd

# Zabbix API credentials
ZABBIX_URL = "https://your-zabbix-url/api_jsonrpc.php"
ZABBIX_USER = "your_username"
ZABBIX_PASSWORD = "your_password"

# Host groups to filter (modify as needed)
HOST_GROUPS = ["Group1", "Group2", "Group3"]

# Connect to Zabbix API
zapi = ZabbixAPI(ZABBIX_URL)
zapi.login(ZABBIX_USER, ZABBIX_PASSWORD)

# Get host groups' IDs
host_groups = zapi.hostgroup.get(filter={"name": HOST_GROUPS}, output=["groupid", "name"])
group_ids = [group["groupid"] for group in host_groups]

# Get hosts in specified host groups
hosts = zapi.host.get(groupids=group_ids, selectTags=["tag", "value"], selectHttpTests=["name", "steps"], output=["hostid", "host"])

data = []
for host in hosts:
    host_id = host["hostid"]
    host_name = host["host"]

    # Get APP tag details
    app_tags = [tag for tag in host.get("tags", []) if tag["tag"] == "APP"]
    
    if not app_tags:
        continue  # Skip if no APP tag found

    tag_name = "APP"
    tag_value = app_tags[0]["value"] if app_tags else ""

    # Get web scenarios
    for scenario in host.get("httpTests", []):
        scenario_name = scenario["name"]
        url = scenario["steps"][0]["url"] if scenario.get("steps") else ""

        data.append([host_name, scenario_name, url, tag_name, tag_value])

# Convert to DataFrame
df = pd.DataFrame(data, columns=["Host Name", "Scenario Name", "URL", "Tag Name", "Tag Value"])

# Save to Excel
output_file = "zabbix_host_scenarios.xlsx"
df.to_excel(output_file, index=False)

print(f"Data saved to {output_file}")

# Logout from Zabbix API
zapi.user.logout()