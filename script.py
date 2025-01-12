# User Interface and Tooltips
# =============================================================================================================================
__title__ = "SECTION FOR MEP ELEMENT"
__author__ = "SUVIR S"
__doc__	 = """This tool create Multiple section for all MEP element like CABLE TRAY,CONDUIT,DUCT,PLUMBING PIPE etc..

This Plugin is Developed by R&D Department Conserve Solution."""

import math
from pyrevit import revit, DB
from pyrevit import script
from pyrevit import forms

output = script.get_output()
app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
selection = uidoc.Selection

# Get the current document and selection
selected_element_ids = selection.GetElementIds()
items = [doc.GetElement(e_id) for e_id in selected_element_ids]

# Check if there are selected items
if not selected_element_ids:
    forms.alert('You must select at least one element to continue the script', exitscript=True)

# Function to get the start and end points of the curve
def GetCurvePoints(curve):
    return curve.GetEndPoint(0), curve.GetEndPoint(1)

# Function to extract the curve from an element
def GetLocation(item):
    if hasattr(item, "Curve"):
        return item.Curve
    if hasattr(item, "Location"):
        loc = item.Location
        if isinstance(loc, DB.LocationCurve):
            return loc.Curve
    return None

# Function to create a section view based on start and end points
def create_section(start_point, end_point, section_type):
    # Vector between start and end points
    direction_vector = end_point - start_point
    midpoint = start_point + 0.5 * direction_vector
    element_direction = direction_vector.Normalize()
    
    # Calculate the true up vector based on the tilt of the element
    up = DB.XYZ.BasisZ if abs(element_direction.Z) < 0.9 else DB.XYZ.BasisY  # Adjust the up vector if the element is mostly vertical
    view_direction = element_direction.CrossProduct(up).Normalize()
    up = view_direction.CrossProduct(element_direction).Normalize()  # Recalculate up to ensure orthogonality
    
    # Define the section bounding box size
    section_offset = 0.5 # Offset for visibility
    y_offset =section_offset+2
    width = direction_vector.GetLength()
    bbox_min = DB.XYZ(-width / 2, -y_offset, -section_offset)
    bbox_max = DB.XYZ(width / 2, y_offset, section_offset)
    
    # Set up the transformation for the section view
    t = DB.Transform.Identity
    t.Origin = midpoint
    t.BasisX = element_direction
    t.BasisY = up
    t.BasisZ = view_direction
    
    # Create the bounding box for the section
    section_box = DB.BoundingBoxXYZ()
    section_box.Transform = t
    section_box.Min = bbox_min
    section_box.Max = bbox_max

    try:
        # Create the section view
        section_view = DB.ViewSection.CreateSection(doc, section_type, section_box)
        if section_view:
            section_view.Scale = 25  # Set the view scale
            section_view.DisplayStyle = DB.DisplayStyle.Wireframe  # Set display style
            section_view.DetailLevel = DB.ViewDetailLevel.Fine  # Set detail level
            section_view.get_Parameter(DB.BuiltInParameter.SECTION_COARSER_SCALE_PULLDOWN_METRIC).Set(1000) # Set the Hide at Scale to 1:1000
            
            output.print_md("Created Section:" + section_view.Name )
        else:
            output.print_md("Section view was not created for the selected Elements.")
    except Exception as e:
        output.print_md("Failed to create section for the selected Elements- " + str(e))

# Function to get the section view family type
def get_section_viewfamily():
    view_family_type_collector = DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType)
    for vft in view_family_type_collector:
        if vft.ViewFamily == DB.ViewFamily.Section:
            return vft.Id
    raise Exception("Section View Family Type not found!")

# Get the section view family type
doc_section_type = get_section_viewfamily()

# Start a transaction for section creation
with revit.Transaction("Create Sections for Selected Elements"):
    for item in items:
        curve = GetLocation(item)
        if curve and isinstance(curve, DB.Curve):
            start_point, end_point = GetCurvePoints(curve)
            create_section(start_point, end_point, doc_section_type)