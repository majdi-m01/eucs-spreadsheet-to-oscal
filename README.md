# Use of Trestle to convert EUCS Spreadsheet to OSCAL

This Python script is designed to convert an EUCS controls spreadsheet into catalogs and profiles based on different severity levels (Basic, Substantial, High) in the (OSCAL) format. The script reads data from an Excel spreadsheet, structures it into an OSCAL-compliant format, and exports catalogs and profiles. Here's a breakdown of the main components and functions of the script:

---

### Script Overview

- **Purpose**: The script reads an EUCS controls spreadsheet, categorizes the controls by severity (Basic, Substantial, High), and outputs the results as OSCAL catalogs and profiles. The output is intended for cybersecurity and information security purposes, providing structured, machine-readable data.
  
- **Input**: An Excel spreadsheet containing EUCS categories, controls, and requirements.
  
- **Output**: JSON files representing an OSCAL catalog and three OSCAL profiles (Basic, Substantial, High).

---

### Modules and Libraries

- **argparse**: Used to handle command-line arguments.
- **datetime**: Provides date and time functionalities.
- **logging**: Manages logging and debug messages.
- **pathlib**: Handles file system paths.
- **sys**: Provides system-related functions.
- **pandas**: Used to read and process the Excel spreadsheet.
- **trestle.oscal.catalog, trestle.oscal.common, trestle.oscal.profile**: Provides the OSCAL data model components (catalogs, profiles, and common elements) used to structure the data.

---

### Functions

#### 1. **`create_catalog`**
Creates and returns an OSCAL catalog object based on a UUID, metadata, groups (categories), and back matter (references/resources).

#### 2. **`create_metadata`**
Generates metadata for the catalog, including properties like keywords, links, roles (publisher, author, contact), and responsible parties. This function simulates an example metadata for the EUCS catalog with ENISA as the organization.

#### 3. **`create_backmatter`**
Creates the back matter of the catalog, including external resources and links (e.g., links to PDF documents).

#### 4. **`create_category`**
Creates a new control group (category) based on a category ID and title. Each group contains its own controls.

#### 5. **`create_control`**
Creates an individual control within a category. The control consists of a control ID, title, properties (such as labels), and additional control parts.

#### 6. **`create_objective_headline`**
Generates the objective section for a control. This is a description of the control’s objective.

#### 7. **`create_requirement_headline`**
Creates a section for control requirements under each control.

#### 8. **`create_requirement`**
Generates a requirement (specific guideline or rule) under a control, based on the severity level (Basic, Substantial, High), the requirement ID, and description.

#### 9. **`write_profile`**
Writes an OSCAL profile to a specified file path. The profile references the catalog and filters controls based on their severity.

#### 10. **`run`**
This is the main function that:
   - Reads the EUCS controls spreadsheet.
   - Processes categories, controls, and requirements.
   - Organizes the controls into different severity levels (Basic, Substantial, High).
   - Generates the catalog and profiles, and writes them to JSON files.

#### 11. **`EUCSConverter`**
This class handles the command-line interface for the script. It initializes arguments (input, output, and EUCS version) and runs the conversion process when invoked.

---

### Main Workflow

1. **Command-line Interface**: The script is run via the command line with arguments specifying the input spreadsheet, the output directory, and the version of the EUCS controls.

2. **Spreadsheet Processing**: The script reads the input Excel file, identifies categories, controls, and requirements based on the contents of specific columns.

3. **Control Categorization**: Controls and their requirements are organized into severity levels (Basic, Substantial, High).

4. **Catalog and Profile Generation**:
   - An OSCAL catalog is created, containing all controls grouped by categories.
   - Separate OSCAL profiles are generated for Basic, Substantial, and High severity levels, each containing a subset of the controls.

5. **File Output**: The resulting catalog and profiles are saved as JSON files in the specified output directory.

---


## Running Demo: Requirements for Excel Sheet:
- The **sheet name** inside the spreadsheet file has to include "controls"
- The **column names** should include:
  *   EUCS Category
  *   EUCS Control
  *   EUCS Requirement
  *   Title
  *   Description
  *   (Profiles) : Basic , Substantial , High
- A **catalogue** row:
  - has an *EUCS Catalogue (ID), Title*
  - has neither an EUCS Control nor an EUCS Requirement. Also, no Description.
- A **control** row:
  - has an *EUCS Catalogue (ID), EUCS Control (ID), Description*
  - not have an EUCS Requirement nor Title.
- A **requirement** row:
  - has an *EUCS Catalogue (ID), EUCS Control (ID), EUCS Requirement (ID), Description*
  - has no Title.

---

## Running Demo: Prepare folders and files

1. Create a new folder name it: "EUCS_controls" insdide "compliance-trestle-demos"
2. Place the python file "create_eucs_catalogs_profiles.py" and the excel sheet "EUCS_Controls_Version_X.XX" inside
3. Create inside "EUCS_controls" another new empty folder, name it: "Outputs"

**Alternatively, you can simply copy paste the folder "EUCS_controls" that I attached.**

**Structure:**

--- EUCS_controls

-------> Outputs

---------------> (*.json* *files* will later show up here)

--------> *create_eucs_catalogs_profiles.py*

--------> *EUCS_Controls_Version_X.XX.xlsx*

---

##  Running Demo: Example Command

```bash
python .\create_eucs_catalogs_profiles.py --input .\EUCS_Controls_Version_1.0.xlsx --output .\Outputs\ --EUCS-version 1.0
```

In this example:
- `--input` specifies the path to the EUCS controls spreadsheet.
- `--output` specifies the directory where the OSCAL catalog and profiles will be saved.
- `--EUCS-version` specifies the version of the EUCS controls.
---

### Expected Outputs
- **Catalog**: A JSON file representing the full EUCS catalog in OSCAL format.
- **Profiles**: Three separate JSON files for each severity level (Basic, Substantial, High), filtering the relevant controls.

