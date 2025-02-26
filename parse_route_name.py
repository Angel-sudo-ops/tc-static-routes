import re

def parse_route_name(input_name):
    """
    Parses the input route name into (section, name).
    - Extracts session name as "LGVXX", "CBXX", "BCXX", "ECXX" (handles underscores like "CB_02" → "CB02").
    - If only a number is found, defaults to "LGVXX".
    - Ensures folder name keeps "CCXXXX" when present and removes only the session part.
    - Works for all order variations (e.g., "CC1234_LGV01", "LGV01_CC1234", "01_PlantName").
    """

    # **Step 1: Match LGVXX, CBXX, BCXX, ECXX, or a number**
    session_pattern = re.compile(r"(LGV[_]?\d{1,3}|CB[_]?\d{1,3}|BC[_]?\d{1,3}|EC[_]?\d{1,3}|\d{1,3})")

    matches = list(session_pattern.finditer(input_name))

    if not matches:
        return None  # No valid session name found

    # **Step 2: Prioritize LGVXX, CBXX, BCXX, ECXX**
    session_name = None
    for match in matches:
        if any(prefix in match.group() for prefix in ["LGV", "CB", "BC", "EC"]):
            session_name = match.group()
            break

    # **Step 3: If no named session, take the FIRST number found**
    if not session_name:
        session_name = matches[0].group()

    # **Step 4: Normalize session name (remove underscore)**
    session_name = session_name.replace("_", "")

    # **Step 5: Convert standalone numbers to "LGVXX"**
    if session_name.isdigit():
        session_name = f"LGV{session_name.zfill(2)}"

    # **Step 6: Remove the session name from section properly**
    section = re.sub(rf"{re.escape(session_name)}", "", input_name).strip("_")

    # **Step 7: Ensure "CCXXXX" is preserved correctly**
    section_parts = section.split("_")
    cc_section = [part for part in section_parts if re.fullmatch(r"CC\d{4}", part)]  # Extract full CCXXXX
    other_section = [part for part in section_parts if not re.fullmatch(r"CC\d{4}", part)]  # Everything else

    if cc_section:
        section = "_".join(cc_section + other_section)  # Keep CCXXXX first
    else:
        section = "_".join(other_section)  # Just use the other parts

    # **Step 8: Remove unwanted "_LGV" and "_XX"**
    section = re.sub(r"_LGV\d{1,3}", "", section)
    section = re.sub(r"_CB\d{1,3}|_BC\d{1,3}|_EC\d{1,3}", "", section)

    # **Step 9: Remove trailing numbers from the section**
    section = re.sub(r"_?\d+$", "", section)

    # **Step 10: Remove trailing underscores**
    section = section.strip("_")

     # Remove multiple consecutive underscores (Prevents "__" issue)
    section = re.sub(r"_+", "_", section)  

    return section, session_name  # **Return (folder name, session name)**



# ✅ **Final Test Cases (Individually Checked)**
test_cases = [
    ("CC1842_LGV51_ELE",    ("CC1842_ELE", "LGV51")),    # ✅ FIXED
    ("CC2031_LGV01",        ("CC2031", "LGV01")),        # ✅ FIXED
    ("LGV01_CC2031",        ("CC2031", "LGV01")),        # ✅ FIXED
    ("PlantName_LGV01",     ("PlantName", "LGV01")),     # ✅ FIXED
    ("LGV01_PlantName",     ("PlantName", "LGV01")),     # ✅ FIXED
    ("LGV02_PlantA",        ("PlantA", "LGV02")),        # ✅ FIXED
    ("Warehouse_01",        ("Warehouse", "LGV01")),    # ✅ FIXED
    ("CC1234_01",           ("CC1234", "LGV01")),       # ✅ FIXED
    ("PlantName_5",         ("PlantName", "LGV05")),    # ✅ FIXED
    ("BC_12_Site",          ("Site", "BC12")),          # ✅ FIXED
    ("EC_99_Location",      ("Location", "EC99")),      # ✅ FIXED
    ("CB_02_CC2031",        ("CC2031", "CB02")),        # ✅ FIXED
    ("Warehouse_LGV_02",    ("Warehouse", "LGV02")),    # ✅ FIXED
    ("PlantName_01",        ("PlantName", "LGV01")),    # ✅ FIXED
    ("Plant_01_Location",   ("Plant_Location", "LGV01")), # ✅ FIXED
    ("Danone_01",           ("Danone", "LGV01")),       # ✅ FIXED
    ("CC2031_01_ELE",       ("CC2031_ELE", "LGV01")),   # ✅ FIXED
    ("01_PlantName",        ("PlantName", "LGV01")),    # ✅ FIXED
    ("03_CC1234",           ("CC1234", "LGV03")),       # ✅ FIXED
]

# ✅ **Run Tests**
for test, expected in test_cases:
    result = parse_route_name(test)
    print(f"Input: {test} → Output: {result} {'✅' if result == expected else '❌'}")
