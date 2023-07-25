def main():
    st.title('IFC to SQL Server and Excel Converter')

    # SQL Server Parameters
    st.sidebar.header("SQL Server Parameters")
    username = st.sidebar.text_input("Username", key="username")
    password = st.sidebar.text_input("Password", key="password", type="password")
    server = st.sidebar.text_input("Server address", key="server")
    database = st.sidebar.text_input("Database name", key="database")
    table_name = st.sidebar.text_input("Table name", key="table_name")

    # Move the file uploader to the sidebar
    uploaded_files = st.file_uploader("Choose IFC files", type='ifc', accept_multiple_files=True, key="file_uploader")

    for uploaded_file in uploaded_files:
        # Create a temporary file
        tfile = tempfile.NamedTemporaryFile(delete=False)
        tfile.write(uploaded_file.getvalue())

        # Convert IFC file to DataFrame
        df = ifc_to_dataframe(tfile.name, uploaded_file.name)  # Pass the file name to ifc_to_dataframe

        # Filter the DataFrame based on the search term and group option
        search_term = st.sidebar.text_input("Enter a search term", key="search_term")
        group_option = st.sidebar.selectbox("Group by", options=["None"] + list(df.columns), key="group_option")
        if search_term:
            df = df[df.apply(lambda row: search_term.lower() in row.to_string().lower(), axis=1)]
        if group_option != "None" and group_option in df.columns:
            df = df.groupby(group_option).first().reset_index()

        # Display the DataFrame
        st.dataframe(df)

        # Export to Excel
        if st.sidebar.button('Export to Excel'):
            st.sidebar.markdown(get_excel_download_link(df), unsafe_allow_html=True)

        # Export to SQL Server
        if st.sidebar.button('Export to SQL Server') and server and username and password and database and table_name:
            connection_string = f'mssql+pyodbc://{username}:{password}@{server}/{database}?driver=ODBC+Driver+17+for+SQL+Server'
            engine = create_engine(connection_string)
            df.to_sql(table_name, engine, if_exists='append')  # append to existing table if it exists
            st.sidebar.write('Data written to SQL Server')

if __name__ == "__main__":
    main()
