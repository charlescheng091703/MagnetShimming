from datetime import datetime
from CdbApiFactory import CdbApiFactory
import csv
import os

factory = CdbApiFactory('https://cdb.aps.anl.gov/cdb')
item_api = factory.getItemApi()
machine_api = factory.getMachineDesignItemApi()

module_ids = []
selectedModule = None

# ModuleDataDirectory = "./ModuleDataDirectory"

        
def FindModulesInCDB(validate):

    # CDB IDs of the Magnet Module Catalog items
    CDBItemID = {}
    CDBItemID['DLMA'] = 110353
    CDBItemID['DLMB'] = 110354
    CDBItemID['FODO'] = 110371
    CDBItemID['QMQA'] = 110369
    CDBItemID['QMQB'] = 110370

    CDBModuleID = {}
#    CDBModuleID["ALL"] = 0
    
    print("   NAME   CDB_ID QR_ID   # Magnets Assigned")
    for magnet_module in CDBItemID.keys():
      for inv_item in item_api.get_items_derived_from_item_by_item_id(CDBItemID[magnet_module]):
        NumAssigned = NumberOfAssignedMagnetQrCodes(inv_item.id)
        print(inv_item.name, inv_item.id, inv_item.qr_id, "       ", NumAssigned)
        if validate == True:
            if NumAssigned == 0:
                continue
        module_ids.append(inv_item.id)
        CDBModuleID[inv_item.name] = inv_item.id
    #print(CDBModuleID)
    return(CDBModuleID)

def NumberOfAssignedMagnetQrCodes(module_id):
    cdbitem = module_id
    item_hierarchyOBJ = item_api.get_item_hierarchy_by_id(cdbitem)
    # Count the number of magnets that have a QR code assigned and are magnets
    # Look for 1) elementName containing a magnet abbreviation (:Q,:S,:M,:F) 
    #          2) only one colon (must eliminate "S{nn}A:M1:STND_US" 
    #          3) has a qr_id assigned
    magnetNames = (":Q", ":S", ":M", ":F")
    numberAssigned = 0
    for item_hierarchy in item_hierarchyOBJ.child_items:
        if item_hierarchy != None:
            elementName = item_hierarchy.derived_element_name
            if any([substring in elementName for substring in magnetNames]) and (elementName.count(':') == 1):
                #print(elementName)
                if item_hierarchy.item != None:
                    if item_hierarchy.item.qr_id:
                        #print(elementName)
                        numberAssigned = numberAssigned + 1
#            if item_hierarchy.item != None:
#                if item_hierarchy.item.qr_id:
             
    return(numberAssigned)
    
def CreateModuleMagnetInfoCsv(module_dict, module_names):   
    # Create .csv files for each module containing magnet info
    # Path to the directory where module specific output files will be stored
    ModuleDataDirectory = "./ModuleDataDirectory"
    if not os.path.exists(ModuleDataDirectory):
        os.makedirs(ModuleDataDirectory)

    #get a list of module IDs from the list of mosule_names 
    module_ids = [module_dict[x] for x in module_names]
    for module_id in module_ids:
        module_item = item_api.get_item_by_id(module_id)
        module_qr_id = module_item.qr_id
        module_name = module_item.name
        print(module_name, "QR_ID=", module_qr_id)

        hierarchy = item_api.get_item_hierarchy_by_id(module_id)
        children = hierarchy.child_items

        # Create a module directory if one does not yet exist
        if not os.path.exists(ModuleDataDirectory + "/" + module_name):
            print("Create directory for ", module_name)
            os.makedirs(ModuleDataDirectory + "/" + module_name)

        OutputFileName = os.path.join(ModuleDataDirectory, module_name, module_name + ".csv")

        module_item_list = []
        print("OutputFile: ", OutputFileName)

        for child in children:
            child_item = child.item
            child_element_name = child.derived_element_name
            if  child_item is None:
              continue

            child_name = child_item.name
            child_qrid = child_item.qr_id
            child_id = child_item.id
            serial_number = child_item.item_identifier1
            if (serial_number == None):
              serial_number = "None"
            catalog_item = child_item.derived_from_item.name
        #    print(child_item.name,child_item.qr_id,child_item.id, child_item.item_identifier1)
        #    print(child_item)
            module_item_list.append([child_name, serial_number, child_element_name, catalog_item, child_qrid])
            print("  ", child_name, serial_number, child_element_name, catalog_item, child_qrid)
        print()
        
        with open(OutputFileName, 'w', newline='') as outfile:
            writer = csv.writer(outfile)
            for row in module_item_list:
                writer.writerow(row);