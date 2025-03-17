import clr
clr.AddReference('RevitServices')
clr.AddReference('RevitNodes')
clr.AddReference('RevitAPI')

from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from Autodesk.Revit.DB import *
if IN[1]:
    # Get the current document (Family Document)
    doc = DocumentManager.Instance.CurrentDBDocument
    
    # Ensure we are inside a Family Document
    if not doc.IsFamilyDocument:
        raise Exception("Please run this script inside a Family document in Revit.")
    
    # Get input from Dynamo (List of values from Excel)
    typeList = IN[0]  # List format [["Type", L, W, H], ...]
    
    # Conversion factor from mm to feet
    MM_TO_FEET = 0.00328084
    
    # Start a Transaction
    TransactionManager.Instance.EnsureInTransaction(doc)
    
    # Get Family Manager
    fam_mgr = doc.FamilyManager
    existing_types = {t.Name: t for t in fam_mgr.Types}  # Store type names with their references
    
    # Retrieve Family Parameters
    param_L = fam_mgr.get_Parameter("L")
    param_W = fam_mgr.get_Parameter("W")
    param_H = fam_mgr.get_Parameter("H")
    
    if not param_L or not param_W or not param_H:
        raise Exception("Parameters 'L', 'W', or 'H' not found in the Family.")
    
    # Read data and create or update family types
    for ins in typeList[1:]:  # Skip headers
        type_name = ins[0]
        length = ins[1] * MM_TO_FEET  # Convert mm to feet
        width = ins[2] * MM_TO_FEET   # Convert mm to feet
        height = ins[3] * MM_TO_FEET  # Convert mm to feet
    
        if type_name in existing_types:
            # If type exists, switch to it and update values
            fam_mgr.CurrentType = existing_types[type_name]
        else:
            # If type does not exist, create it
            new_type = fam_mgr.NewType(type_name)
            existing_types[type_name] = new_type
    
        # Set parameter values
        fam_mgr.Set(param_L, length)
        fam_mgr.Set(param_W, width)
        fam_mgr.Set(param_H, height)
    
    # Commit transaction
    TransactionManager.Instance.TransactionTaskDone()
    
    # Output success message
    OUT = "Family types created/updated successfully!"
