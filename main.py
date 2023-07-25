import streamlit as st
import pandas as pd
import ifcopenshell
import tempfile
from sqlalchemy import create_engine
import base64
import io
import os



def ifc_to_dataframe(ifc_file_path, file_name):
    """
    Convert an IFC file to a DataFrame.

    Args:
        ifc_file_path (str): The path to the IFC file.
        file_name (str): The name of the IFC file.

    Returns:
        pd.DataFrame: The DataFrame containing the IFC data.
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

    # Convert data to DataFrame
    df = pd.DataFrame(data)

    # Delete rows where Type equals 'IfcGroup' and 'IfcOpeningElement'
    df = df[df.Type != 'IfcGroup']
    df = df[df.Type != 'IfcOpeningElement']

    return df


def get_excel_download_link(df, filename="data.xlsx"):
    """
    Generates a link allowing the data in a given panda dataframe to be downloaded
    in:  dataframe
    out: href string
    """
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

    if uploaded_files:
        for uploaded_file in uploaded_files:
            # Create a temporary file
            tfile = tempfile.NamedTemporaryFile(delete=False)
            tfile.write(uploaded_file.getvalue())

            # Convert IFC file to DataFrame
            df = ifc_to_dataframe(tfile.name, uploaded_file.name)

            # Filter the DataFrame based on the search term and group option
            search_key = f"search_term_{uploaded_file.name}"
            search_term = st.sidebar.text_input("Enter a search term", key=search_key)

            # ... (previous code)

            # Export to Excel
            if st.sidebar.button('Export to Excel'):
                st.sidebar.markdown(get_excel_download_link(df), unsafe_allow_html=True)

            # Export to SQL Server
            if st.sidebar.button('Export to SQL Server') and server and username and password and database and table_name:
                connection_string = f'mssql+pyodbc://{username}:{password}@{server}/{database}?driver=ODBC+Driver+17+for+SQL+Server'
                engine = create_engine(connection_string)
                df.to_sql(table_name, engine, if_exists='append')  # append to existing table if it exists
                st.sidebar.write('Data written to SQL Server')

            # Display the DataFrame
            st.dataframe(df)

if __name__ == "__main__":
    main()