"""Script to convert EUCS controls spreadsheet into catalogs and profiles based on severity (Basic, Substantial, High)."""

import argparse
import datetime
import logging
import pathlib
import sys
from typing import List, Optional
from uuid import uuid4

from ilcli import Command  # CLI wrapper utility
import pandas as pd  # Library to handle Excel file operations

# Importing necessary modules from the OSCAL data model (Trestle)
import trestle.oscal.catalog as oscat
import trestle.oscal.common as oscommon
import trestle.oscal.profile as ospro

# Set up the logger for debugging and error messages
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
logger.addHandler(logging.StreamHandler(sys.stdout))

# Function to create an OSCAL Catalog object with metadata, groups, and back matter
def create_catalog(uuid, metadata, groups, back_matter):
    return oscat.Catalog(
        uuid=uuid, 
        metadata=metadata, 
        groups=groups, 
        back_matter=back_matter
    )

# Function to create metadata for the catalog
def create_metadata():
    properties = []
    # Add keywords related to cybersecurity and information security
    properties.append(oscommon.Property(
        name='keywords', 
        value='cybersecurity, information security, information system, OSCAL, Open Security Controls Assessment Language'
    ))

    # Create links for the metadata
    links = []
    links.append(oscommon.Link(rel="alternate", href="3ab41d8a-66e3-4732-ae28-07405dad5127"))

    # Create roles (e.g., publisher, author, contact)
    roles = []
    roles.append(oscommon.Role(id='publisher', title='Source document converter to OSCAL.'))
    roles.append(oscommon.Role(id='author', title='Source document author.'))
    roles.append(oscommon.Role(id='contact', title='Contact.'))

    # Add email addresses and physical addresses for the parties
    email_addresses = []
    email_addresses.append(oscommon.EmailAddress(__root__="name@domain.eu"))
    address_lines = []
    address_lines.append("ENISA")
    address_lines.append("Attn: Somebody")
    address_lines.append("1 Some Street")
    addresses = []
    addresses.append(oscommon.Address(addr_lines=address_lines, city="City", country="EU"))
    
    # Create party information (organization)
    parties = []
    parties.append(oscommon.Party(
        type='organization', 
        uuid="f550d94e-0f01-415f-a1e2-c8188c9ff4a5", 
        name="ENISA", 
        email_addresses=email_addresses, 
        addresses=addresses
    ))

    # Set responsible parties for publisher, author, and contact roles
    partyuuids = []
    partyuuids.append("f550d94e-0f01-415f-a1e2-c8188c9ff4a5")
    responsible_parties = []
    responsible_parties.append(oscommon.ResponsibleParty(role_id="publisher", party_uuids=partyuuids))
    responsible_parties.append(oscommon.ResponsibleParty(role_id="author", party_uuids=partyuuids))
    responsible_parties.append(oscommon.ResponsibleParty(role_id="contact", party_uuids=partyuuids))
    
    # Return the metadata object containing title, publication dates, roles, parties, etc.
    return oscommon.Metadata(
        title="Sample EUCS Catalog", 
        published=datetime.datetime.now().astimezone(), 
        last_modified=datetime.datetime.now().astimezone(), 
        version="1.0", 
        oscal_version="1.1.2", 
        props=properties, 
        links=links, 
        roles=roles, 
        parties=parties, 
        responsible_parties=responsible_parties, 
        remarks="The following is a short excerpt from EUCS Catalog. This work is provided here under copyright fair use for non-profit, educational purposes only. Copyrights for this work are held by the publisher."
    )

# Function to create the back matter (resources and references) of the catalog
def create_backmatter():
    rlinks = []
    # Add a reference link to the EUCS PDF document
    rlinks.append(oscommon.Rlink(media_type="application/pdf", href="https://enisa.europa.eu/publications/eucs.pdf"))

    resources = []
    # Create a resource object that links to the EUCS PDF
    resources.append(oscommon.Resource(uuid="3ab41d8a-66e3-4732-ae28-07405dad5127", title="EUCS prCEN/TS (PDF)", rlinks=rlinks))

    return oscommon.BackMatter(
        resources=resources
    )

# Function to create a new category (group) of controls in the catalog
def create_category(category_id, category_title):
    properties = []
    # Set the label for the category
    properties.append(oscommon.Property(name='label', value=category_id + '.'))

    return oscat.Group(
        id='eucs-' + category_id, 
        title=category_title, 
        props=properties, 
        controls=[]  # Initialize with an empty list of controls
    )

# Function to create a new control within a category
def create_control(category_id, control_id, control_title):
    properties = []
    # Set the label for the control
    properties.append(oscommon.Property(name='label', value=control_id[-1] + '.'))

    return oscat.Control(
        id='eucs-' + category_id + '.' + control_title.split()[0],  # Unique ID for the control
        class_='eucs',  # Class name (EUCS control)
        title=control_title,  # Title of the control
        props=properties, 
        parts=[]  # Initialize with an empty list of control parts
    )

# Function to create the objective headline (the goal of the control)
def create_objective_headline(category_id, control_title, control_objective):
    properties = []
    properties.append(oscommon.Property(name='label', value='1.'))  # Label for the objective
    properties.append(oscommon.Property(name='alt-identifier', value='Objective'))  # Alternative identifier for the objective part

    return oscommon.Part(
        name='statement',  # The part's name (statement of the objective)
        id='eucs-' + category_id + '.' + control_title.split()[0] + '_obj',  # Unique ID for the objective
        props=properties, 
        prose=control_objective  # The actual content of the objective
    )

# Function to create the requirements headline (the specific actions to fulfill the control)
def create_requirement_headline(category_id, control_title):
    properties = []
    properties.append(oscommon.Property(name='label', value='2.'))  # Label for the requirements
    properties.append(oscommon.Property(name='alt-identifier', value='Requirements'))  # Alternative identifier for the requirements part

    return oscommon.Part(
        name='statement',  # The part's name (statement of requirements)
        id='eucs-' + category_id + '.' + control_title.split()[0] + '_req',  # Unique ID for the requirements
        props=properties, 
        parts=[]  # Initialize with an empty list of requirement parts
    )

# Function to create a requirement within the control (specific rule or action)
def create_requirement(severity_level, requirement_id, category_id, control_title, requirement_description):
    properties = []
    # Set an alternative identifier for the requirement part, based on the severity level (Basic, Substantial, High)
    properties.append(oscommon.Property(name='alt-identifier', class_=severity_level, value=requirement_id))

    return oscommon.Part(
        name='item',  # Name for this requirement
        class_=severity_level,  # Class based on the severity level
        id='eucs-' + category_id + '.' + control_title.split()[0] + '_req' + '.' + requirement_id[-2:],  # Unique ID for the requirement
        props=properties, 
        prose= requirement_id + ' - ' + requirement_description  # The content of the requirement
    )

# Function to write the OSCAL profile file, filtering the control list based on severity
def write_profile(profile: ospro.Profile, control_list: List[str], path: pathlib.Path):
    """Fill in control list and write the profile."""
    include_controls: List[str] = []
    selector = ospro.SelectControl()
    selector.with_ids = control_list  # Add controls to include based on severity level
    include_controls.append(selector)
    profile.imports[0].include_controls = include_controls

    # Write the profile to the specified path
    profile.oscal_write(path)

import pathlib
import pandas as pd
import argparse
import sys

def run(input_xls: pathlib.Path, output_directory: pathlib.Path, EUCS_version: str):
    """
    Main function to process an Excel file and convert the controls and requirements 
    into a catalog and profiles for different severity levels.
    
    Parameters:
    input_xls (pathlib.Path): The path to the input Excel file.
    output_directory (pathlib.Path): The directory where the output files will be saved.
    EUCS_version (str): The version number of the EUCS catalog.
    """
    
    # Load the Excel file
    excel_handler = pd.ExcelFile(input_xls)
    df = None
    
    # Search for the sheet containing 'controls' in its name
    for key in excel_handler.sheet_names:
        if 'controls' in str(key).lower():
            sheet_name = key
    
    # Load the selected sheet into a DataFrame
    df = pd.read_excel(input_xls, sheet_name=sheet_name, header=0, dtype=str)

    # Define key column names for the data
    category_key = 'EUCS Category'
    control_key = 'EUCS Control'
    requirement_key = 'EUCS Requirement'
    title_key = 'Title'
    description_key = 'Description'
    basic_key = 'Basic'
    substantial_key = 'Substantial'	
    high_key = 'High'
    
    # Lists to store the requirements for different profiles
    basic_profile = []
    substantial_profile = []
    high_profile = []

    # Create metadata and backmatter objects for the catalog
    catalog_metadata = create_metadata()
    backmatter = create_backmatter()

    # Create a catalog object with initial metadata and an empty list of groups (categories)
    catalog = create_catalog(uuid="74c8ba1e-5cd4-4ad1-bbfd-d888e2f6c724", metadata=catalog_metadata, groups=[], back_matter=backmatter)
    
    current_loop_category = None  # Track the current category
    current_loop_control = None   # Track the current control
    
    # Iterate through each row in the DataFrame
    for _, row in df.iterrows():        
        # Identify categories, controls, and requirements based on the conditions provided
        new_loop_category = row[category_key]
        new_loop_control = row[control_key]

        # Adding a new category if the row contains a category but no control or requirement
        if pd.isna(row[control_key]) and pd.isna(row[requirement_key]) and pd.notna(row[category_key]):
            # Add new category to catalog if it's a new one
            if new_loop_category != current_loop_category or current_loop_category is None:
                category = create_category(
                    category_id=row[category_key], 
                    category_title=row[title_key]
                )
                catalog.groups.append(category)

        # Adding a new control if the row contains a control but no requirement
        elif pd.isna(row[requirement_key]) and pd.notna(row[category_key]) and pd.notna(row[control_key]):
            # Add new control to current category if it's a new one
            if new_loop_control != current_loop_control or current_loop_control is None:
                control = create_control(
                    category_id=row[category_key], 
                    control_id=row[control_key], 
                    control_title=row[title_key]
                )
                category.controls.append(control)

                # Add an objective headline part to the control
                objective_headline = create_objective_headline(
                    category_id=row[category_key], 
                    control_title=row[title_key], 
                    control_objective=row[description_key]
                )
                control.parts.append(objective_headline)
                
                # Add a requirement headline part to the control
                requirement_headline = create_requirement_headline(
                    category_id=row[category_key], 
                    control_title=row[title_key]
                )
                control.parts.append(requirement_headline)
                current_loop_control_title = row[title_key]            
        
        # Adding requirements when all keys (category, control, requirement) are present
        elif pd.notna(row[category_key]) and pd.notna(row[control_key]) and pd.notna(row[requirement_key]):
            
            # Add a basic requirement if present
            if pd.notna(row[basic_key]):
                requirement = create_requirement(
                    severity_level='basic', 
                    requirement_id=row[requirement_key], 
                    category_id=row[category_key], 
                    control_title=current_loop_control_title, 
                    requirement_description=row[description_key]
                )
                # Add to basic profile
                basic_profile.append(row[requirement_key])

                # Append the requirement to the control's parts
                requirement_headline.parts.append(requirement)
                
            # Add a substantial requirement if present
            if pd.notna(row[substantial_key]):
                # If the requirement ends with 'B', modify the key to 'S' for substantial
                if row[requirement_key][-1] == 'B':
                    new_requirement_key = row[requirement_key][:-1] + 'S'
                    requirement = create_requirement(
                        severity_level='substantial', 
                        requirement_id=new_requirement_key, 
                        category_id=row[category_key], 
                        control_title=current_loop_control_title, 
                        requirement_description=row[description_key]
                    )
                    substantial_profile.append(new_requirement_key)
                else:
                    requirement = create_requirement(
                        severity_level='substantial', 
                        requirement_id=row[requirement_key], 
                        category_id=row[category_key], 
                        control_title=current_loop_control_title, 
                        requirement_description=row[description_key]
                    )
                    substantial_profile.append(row[requirement_key])
                
                # Append the requirement to the control's parts
                requirement_headline.parts.append(requirement)

            # Add a high requirement if present
            if pd.notna(row[high_key]):
                # Modify the key if the requirement ends with 'B' or 'S' to form 'H' for high severity
                if row[requirement_key][-1] == 'B' or row[requirement_key][-1] == 'S':
                    new_requirement_key = row[requirement_key][:-1] + 'H'
                    requirement = create_requirement(
                        severity_level='high', 
                        requirement_id=new_requirement_key, 
                        category_id=row[category_key], 
                        control_title=current_loop_control_title, 
                        requirement_description=row[description_key]
                    )
                    high_profile.append(new_requirement_key)
                else:
                    requirement = create_requirement(
                        severity_level='high', 
                        requirement_id=row[requirement_key], 
                        category_id=row[category_key], 
                        control_title=current_loop_control_title, 
                        requirement_description=row[description_key]
                    )
                    high_profile.append(row[requirement_key])
                
                # Append the requirement to the control's parts
                requirement_headline.parts.append(requirement)
        
        # Update the current loop category and control for the next iteration
        current_loop_category = row[category_key]
        current_loop_control = row[control_key]

    # Define the path for saving the catalog as a JSON file
    catalog_path: pathlib.Path = pathlib.Path.cwd() / output_directory / f'EUCS_controls_version_{EUCS_version}_catalog.json'
    catalog.oscal_write(catalog_path)

    # Create a profile for each severity level (Basic, Substantial, High)
    profile = ospro.Profile(
        uuid="74c8ba1e-5cd4-4ad1-bbfd-d888e2f6c724",
        metadata=catalog_metadata,
        imports=[ospro.Import(href=catalog_path.name)]
    )

    # Save the Basic profile
    profile_path: pathlib.Path = output_directory / f'EUCS_version_{EUCS_version}_profile_Basic.json'
    write_profile(profile, basic_profile, profile_path)

    # Save the Substantial profile
    profile_path: pathlib.Path = output_directory / f'EUCS_version_{EUCS_version}_profile_Substantial.json'
    write_profile(profile, substantial_profile, profile_path)

    # Save the High profile
    profile_path: pathlib.Path = output_directory / f'EUCS_version_{EUCS_version}_profile_High.json'
    write_profile(profile, high_profile, profile_path)


class EUCSConverter(Command):
    """Converter CLI wrapper class for handling command-line arguments."""
    
    def _init_arguments(self) -> None:
        """
        Initialize command-line arguments.
        """
        self.add_argument('-i', '--input', help='Path to the input Excel file.', type=pathlib.Path, required=True)
        self.add_argument('-o', '--output', help='Path to the output directory.', type=pathlib.Path, default=pathlib.Path.cwd())
        self.add_argument('-v', '--EUCS-version', help='Version of the EUCS catalog.', type=str, default="1.0")

    def _run(self, args: argparse.Namespace) -> int:
        """
        Execute the run function using the provided arguments.
        
        Parameters:
        args (argparse.Namespace): Parsed command-line arguments.
        
        Returns:
        int: Exit code (0 for success).
        """
        run(args.input, args.output, args.EUCS_version)


# Entry point for the script
if __name__ == '__main__':
    sys.exit(EUCSConverter().run())

