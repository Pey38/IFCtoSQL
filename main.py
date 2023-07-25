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
    st.title('IFC to SQL Server and excel Converter')

    # Move the file uploader to the sidebar
    uploaded_files = st.sidebar.file_uploader("Choose IFC files", type='ifc', accept_multiple_files=True, key="file_uploader")

    # Create a dictionary to store each DataFrame
    dfs = {}

    for uploaded_file in uploaded_files:
        # Create a temporary file
        tfile = tempfile.NamedTemporaryFile(delete=False)
        tfile.write(uploaded_file.getvalue())

        # Convert IFC file to DataFrame and store it in the dictionary
        df = ifc_to_dataframe(tfile.name, uploaded_file.name)  # Pass the file name to ifc_to_dataframe
        dfs[uploaded_file.name] = df

    # Create a sidebar for the search and group options
    st.sidebar.header("Options")

    # Add a checkbox for the user to select each model
    model_options = {model: st.sidebar.checkbox("Select " + model, value=True) for model in dfs.keys()}

    # Add a text input for the search term in the sidebar
    search_term = st.sidebar.text_input("Enter a search term", key="search_term")

    # Add a select box for the group option in the sidebar
    group_option = st.sidebar.selectbox("Group by", options=["None"] + list(set(col for df in dfs.values() for col in df.columns)), key="group_option")

    # Create a list to store the dataframes of the selected models
    selected_dfs = []

    # For each selected model, add its DataFrame to the list
    for model_option, is_selected in model_options.items():
        if is_selected:
            # Get the DataFrame for the chosen model
            df = dfs[model_option]

            # Filter the DataFrame based on the search term
            if search_term:
                df = df[df.apply(lambda row: search_term.lower() in row.to_string().lower(), axis=1)]

            # Group the DataFrame based on the group option
            if group_option != "None" and group_option in df.columns:
                df = df.groupby(group_option).first().reset_index()

            # Add the DataFrame to the list
            selected_dfs.append(df)

    # Concatenate all the selected dataframes into a single dataframe
    if selected_dfs:
        all_df = pd.concat(selected_dfs, ignore_index=True)

        # Display the DataFrame
        st.dataframe(all_df)
        
          # Export to Excel button
        if st.sidebar.button('Export to Excel'):
            st.sidebar.markdown(get_excel_download_link(all_df), unsafe_allow_html=True)
            
           # Export to SQL Server button
        if st.sidebar.button('Export to SQL Server') and server and username and password and database and table_name:
            connection_string = f'mssql+pyodbc://{username}:{password}@{server}/{database}?driver=ODBC+Driver+17+for+SQL+Server'
            engine = create_engine(connection_string)
            all_df.to_sql(table_name, engine, if_exists='replace')  # replace existing table if it exists
            st.sidebar.write('Data written to SQL Server')

      

        # SQL Server Parameters
        st.sidebar.header("SQL Server Parameters")
        username = st.sidebar.text_input("Username", key="username")
        password = st.sidebar.text_input("Password", key="password", type="password")
        server = st.sidebar.text_input("Server address", key="server")
        database = st.sidebar.text_input("Database name", key="database")
        table_name = st.sidebar.text_input("Table name", key="table_name")

     

if __name__ == "__main__":
    main()
    