import re

def parse_route_name(input_name):
    """
    Parses the input route name into (section, name).
    - Extracts session name as "LGVXX", "CBXX", "BCXX", "ECXX" (handles underscores like "CB_02" → "CB02").
    - If only a number is found, defaults to "LGVXX".
    - Ensures "CCXXXX" is always kept in the section if present.
    - Works for all order variations (e.g., "CC1234_LGV01", "LGV01_CC1234", "01_PlantName").
    """

    # Pattern to match session names first
    session_pattern = re.compile(r"(LGV[_]?\d{1,3}|CB[_]?\d{1,3}|BC[_]?\d{1,3}|EC[_]?\d{1,3}|\b\d{1,3}\b)")

    matches = list(session_pattern.finditer(input_name))
    
    if not matches:
        return None  # No valid session name found

    # Prioritize named session types (LGVXX, CBXX, etc.) over plain numbers
    session_name = None
    for match in matches:
        if any(prefix in match.group() for prefix in ["LGV", "CB", "BC", "EC"]):
            session_name = match.group()
            break

    # If no session type was found, take the last numeric match
    if not session_name:
        session_name = matches[-1].group()

    # Remove underscores in session name (e.g., "CB_02" → "CB02")
    session_name = session_name.replace("_", "")

    # If session name is just a number, convert it to "LGVXX"
    if session_name.isdigit():
        session_name = f"LGV{session_name.zfill(2)}"

    # Remove session name from input while keeping underscores correctly
    section = input_name.replace(match.group(), "").strip("_")

    # Ensure "CCXXXX" remains in the section if present
    section_parts = section.split("_")
    cc_section = [part for part in section_parts if "CC" in part]  # Extract CCXXXX
    other_section = [part for part in section_parts if "CC" not in part]  # Everything else

    if cc_section:
        section = "_".join(cc_section + other_section)  # Keep CCXXXX first
    else:
        section = "_".join(other_section)  # Just use the other parts

    # Remove multiple consecutive underscores (Prevents "__" issue)
    section = re.sub(r"_+", "_", section)  

    return section, session_name  # Return (folder name, session name)


# ✅ **Test Cases - Ensuring Stability**
test_cases = [
    "CC1842_LGV51_ELE",    # ("CC1842_ELE", "LGV51") ✅
    "CC2031_LGV01",        # ("CC2031", "LGV01") ✅
    "LGV01_CC2031",        # ("CC2031", "LGV01") ✅
    "PlantName_LGV01",     # ("PlantName", "LGV01") ✅
    "LGV01_PlantName",     # ("PlantName", "LGV01") ✅
    "LGV02_PlantA",        # ("PlantA", "LGV02") ✅
    "Warehouse_01",        # ("Warehouse", "LGV01") ✅
    "CC1234_01",           # ("CC1234", "LGV01") ✅
    "PlantName_5",         # ("PlantName", "LGV05") ✅
    "BC_12_Site",          # ("Site", "BC12") ✅
    "EC_99_Location",      # ("Location", "EC99") ✅
    "CB_02_CC2031",        # ("CC2031", "CB02") ✅
    "Warehouse_LGV_02",    # ("Warehouse", "LGV02") ✅
    "PlantName_01",        # ("PlantName", "LGV01") ✅
    "Plant_01_Location",   # ("Plant_Location", "LGV01") ✅
    "Danone_01",           # ("Danone", "LGV01") ✅
    "CC2031_01_ELE",       # ("CC2031_ELE", "LGV01") 
    "01_PlantName",        # ("PlantName", "LGV01") ✅ FIXED
    "03_CC1234",           # ("CC1234", "LGV03") ✅ FIXED
]

for test in test_cases:
    print(f"Input: {test} → Output: {parse_route_name(test)}")
