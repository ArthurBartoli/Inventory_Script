def scan_export(json_data):
    '''
    This function aims at reading the json export and output a dictionary of all metadata sorted
    by reports, datasets, users, workspaces, etc... 
    '''
    # Set the result storage
    keys = ["search", "reports", "workspaces", "datasets", "users", "dataflows", "dashboards"]
    res_final = {k: {"PersonalWorkspace": {}, "SharedWorkspace": {}} for k in keys}
    is_personal_workspace = lambda name: "PersonalWorkspace" in name or "My workspace" in name
    
    ### Search the workspaces
    for t in range(len(json_data)):
        workspaceId = str(json_data[t]["Id"])
        workspaceName = str(json_data[t]["Name"])
        type_espace = "PersonalWorkspace" if is_personal_workspace(workspaceName) else "SharedWorkspace"
        
        res_final["workspaces"][type_espace][workspaceId] = {'workspaceName': workspaceName}
        res_final["search"][type_espace][workspaceId] = {'workspaceName': workspaceName}

        # Set data for browsing
        keys = ["Reports", "Dashboards", "Datasets", "Users", "Dataflows"]
        data = {key: json_data[t][key] for key in keys}
        
        ## We browse through the REPORTS
        res_final["search"][type_espace][workspaceId]["reports"] = {}
        if data["Reports"]:
            for report in data["Reports"]:
                # We keep the name for the entry one level above in the data tree
                tmpName = report["Id"]
                del report["Id"]
                
                # We create an entry for each report
                res_final["search"][type_espace][workspaceId]["reports"][tmpName] = report
                if tmpName not in res_final["reports"][type_espace]:
                    res_final["reports"][type_espace][tmpName] = report
        
        ## We browse through the DASHBOARDS            
        res_final["search"][type_espace][workspaceId]["dashboards"] = {}
        if data["Dashboards"]:
            for dashboard in data["Dashboards"]:
                # We keep the name for the entry
                tmpName = dashboard["Id"]
                del dashboard["Id"]
                
                # We create an entry for each report
                res_final["search"][type_espace][workspaceId]["dashboards"][tmpName] = dashboard
                if tmpName not in res_final["dashboards"][type_espace]:
                    res_final["dashboards"][type_espace][tmpName] = dashboard
                    
        ## We search through the DATASETS   
        res_final["search"][type_espace][workspaceId]["datasets"] = {}
        if data["Datasets"]:
            for dataset in data["Datasets"]:
                # We keep the name for the entry
                tmpName = dataset["Id"]
                del dataset["Id"]
                
                # We create an entry for each report
                res_final["search"][type_espace][workspaceId]["datasets"][tmpName] = dataset
                if tmpName not in res_final["datasets"][type_espace]:
                    res_final["datasets"][type_espace][tmpName] = dataset
        
        ## We search through the USERS        
        res_final["search"][type_espace][workspaceId]["users"] = {}
        if data["Users"] :
            for user in data["Users"]:
                # We keep the name for the entry
                tmpName = user["UserPrincipalName"]
                del user["UserPrincipalName"]
                
                # We create an entry for each report
                res_final["search"][type_espace][workspaceId]["users"][tmpName] = user
                if tmpName not in res_final["users"][type_espace]:
                    del user["AccessRight"]
                    res_final["users"][type_espace][tmpName] = user
        
        ## We search through the dataflows        
        res_final["search"][type_espace][workspaceId]["dataflows"] = {}
        if data["Dataflows"] :
            for dataflow in data["Dataflows"]:
                # We keep the name for the entry
                tmpName = dataflow["Id"]
                del dataflow["Id"]
                
                # We create an entry for each report
                res_final["search"][type_espace][workspaceId]["dataflows"][tmpName] = dataflow
                if tmpName not in res_final["dataflows"][type_espace]:
                    res_final["dataflows"][type_espace][tmpName] = dataflow

    return res_final
                        
