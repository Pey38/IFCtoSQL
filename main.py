import streamlit as st
import ifcopenshell
import tempfile
from sqlalchemy import create_engine
import base64
import io
import os


def ifc_to_list(ifc_file_path, file_name):
    """
    Convert an IFC file to a list of dictionaries.

    Args:
        ifc_file_path (str): The path to the IFC file.
        file_name (str): The name of the IFC file.

    Returns:
        list: The list containing dictionaries with the IFC data.
    """

    # Use ifcopenshell or similar library to read the IFC file
    ifc = ifcopenshell.open(ifc_file_path)

    # Extract all data from the IFC file
    data = []
    for entity in ifc.by_type('IfcObject'):
        entity_data = {
            "GlobalId": entity.GlobalId,
            "Name": entity.Name,
            "Type": entity.is_a(),
            "FileName": file_name,
        }

        # Iterate over IsDefinedBy relationships, which may include property sets
        for rel in entity.IsDefinedBy:
            # Check if this relationship is a property set
            if rel.is_a('IfcRelDefinesByProperties'):
                # Get the property set definition
                property_definition = rel.RelatingPropertyDefinition
                # Check if property_definition is a property set
                if property_definition.is_a('IfcPropertySet'):
                    # Iterate over the properties in the set
                    for prop in property_definition.HasProperties:
                        # Check if the property has a simple value (not all do)
                        if prop.is_a('IfcPropertySingleValue'):
                            # Add the property value to entity_data
                            entity_data[prop.Name] = prop.NominalValue.wrappedValue if prop.NominalValue else None

        data.append(entity_data)

    return data


def get_excel_download_link(data_list, filename="data.xlsx"):
    """
    Generates a link allowing the data in a given list of dictionaries to be downloaded in an Excel file.

    Args:
        data_list (list): The list of dictionaries containing the data.
        filename (str): The name of the Excel file to be downloaded.

    Returns:
        str: The HTML link for downloading the Excel file.
    """
    df = pd.DataFrame(data_list)
    excel_file = io.BytesIO()
    df.to_excel(excel_file, index=False)
    excel_file.seek(0)
    b64 = base64.b64encode(excel_file.read()).decode()  # some strings
    return f'<a href="data:application/octet-stream;base64,{b64}" download="{filename}">Download excel file</a>'


def main():
    st.title('IFC to SQL Server and Excel Converter')

    # SQL Server Parameters
    st.sidebar.header("SQL Server Parameters")
    username = st.sidebar.text_input("Username")
    password = st.sidebar.text_input("Password", type="password")
    server = st.sidebar.text_input("Server address")
    database = st.sidebar.text_input("Database name")
    table_name = st.sidebar.text_input("Table name")

    # File Uploader
    uploaded_files = st.file_uploader("Choose IFC files", type='ifc', accept_multiple_files=True)

    # Initialize an empty list to gather data from all files
    all_data_list = []

    if uploaded_files:
        for uploaded_file in uploaded_files:
            # Create a temporary file
            tfile = tempfile.NamedTemporaryFile(delete=False)
            tfile.write(uploaded_file.getvalue())

            # Convert IFC file to a list of dictionaries
            data_list = ifc_to_list(tfile.name, uploaded_file.name)

            # Append data from this file to the main list
            all_data_list.extend(data_list)

    if all_data_list:
        # Filter the data based on the search term and group option
        search_term = st.sidebar.text_input("Enter a search term")
        group_option = st.sidebar.selectbox("Group by", options=["None", "GlobalId", "Name", "Type"])

        if search_term:
            all_data_list = [item for item in all_data_list if search_term.lower() in str(item).lower()]

        if group_option != "None" and group_option in all_data_list[0]:
            # Group the data based on the selected column
            grouped_data = {}
            for item in all_data_list:
                key = item[group_option]
                if key not in grouped_data:
                    grouped_data[key] = item
            all_data_list = list(grouped_data.values())

        # Display the data
        st.write(all_data_list)

        # Export to Excel
        if st.sidebar.button('Export to Excel'):
            st.sidebar.markdown(get_excel_download_link(all_data_list), unsafe_allow_html=True)

        # Export to SQL Server
        if st.sidebar.button('Export to SQL Server') and server and username and password and database and table_name:
            connection_string = f'mssql+pyodbc://{username}:{password}@{server}/{database}?driver=ODBC+Driver+17+for+SQL+Server'
            engine = create_engine(connection_string)
            with engine.connect() as conn:
                for item in all_data_list:
                    conn.execute(f"INSERT INTO {table_name} VALUES {tuple(item.values())}")
            st.sidebar.write('Data written to SQL Server')
    else:
        st.write("No data found. Please upload IFC files.")


if __name__ == "__main__":
    main()
