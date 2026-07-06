import streamlit as st
import pandas as pd
import io

st.title("GHG Target Sheet Generator")

st.markdown("Upload your source Excel file. This app will calculate the required target dataframes and allow you to download an updated Excel file with the two new dataframes appended as separate sheets.")

uploaded_file = st.file_uploader("Upload Excel File (must contain source data in the first sheet)", type=["xlsx", "xls"])

if uploaded_file is not None:
    st.write("Loading data...")
    try:
        df = pd.read_excel(uploaded_file, sheet_name=0)
    except Exception as e:
        st.error(f"Error loading file: {e}")
        st.stop()

    needed_cols = [
        'LevelGOrgPlantName', 'ScopeDriverName', 'ScopeName', 
        'ActivityDataValue', 'CO2e_Value', 'IsProduct', 
        'IsToBeSubstracted', 'IsAccepted'
    ]
    missing = [c for c in needed_cols if c not in df.columns]
    
    if missing:
        st.error(f"Error: Missing required columns in the source sheet: {missing}")
        st.stop()

    # ---------------------------------------------------------
    # TABLE 1: Absolute Emission, Production, Intensity
    # ---------------------------------------------------------
    c2_pos = df[(df['IsAccepted'] == 1) & (df['IsToBeSubstracted'] == 0)]['CO2e_Value'].sum()
    c2_neg = df[(df['IsAccepted'] == 1) & (df['IsToBeSubstracted'] == 1)]['CO2e_Value'].sum()
    absolute_emission = c2_pos - c2_neg
    
    production = df[df['IsProduct'] == 1]['ActivityDataValue'].sum()
    intensity = absolute_emission / production if production != 0 else 0
    
    df_table1 = pd.DataFrame({
        'Metric': ['Absolute Emission', 'Production', 'Intensity'],
        'Value': [absolute_emission, production, intensity]
    })

    st.subheader("Table 1: Target Overview")
    st.dataframe(df_table1)

    # ---------------------------------------------------------
    # TABLE 2: Grouped by Plant and Scope Driver
    # ---------------------------------------------------------
    groups = df.groupby(['LevelGOrgPlantName', 'ScopeDriverName'])
    
    results = []
    for (plant, driver), group in groups:
        ad_val = group[group['ScopeName'] == 'Scope 1']['ActivityDataValue'].sum()
        
        def calc_emission(scope):
            g_scope = group[group['ScopeName'] == scope]
            pos = g_scope[g_scope['IsToBeSubstracted'] == 0]['CO2e_Value'].sum()
            neg = g_scope[g_scope['IsToBeSubstracted'] == 1]['CO2e_Value'].sum()
            return pos - neg
            
        scope1_em = calc_emission('Scope 1')
        scope2_em = calc_emission('Scope 2')
        scope3_em = calc_emission('Scope 3')
        total_em = scope1_em + scope2_em + scope3_em
        
        t_tcs = ad_val / production if production != 0 else 0
        scope1_int = scope1_em / production if production != 0 else 0
        scope2_int = scope2_em / production if production != 0 else 0
        scope3_int = scope3_em / production if production != 0 else 0
        total_int = total_em / production if production != 0 else 0
        
        results.append({
            'LevelGOrgPlantName': plant,
            'ScopeDriverName': driver,
            'AD': ad_val,
            't/tcs': t_tcs,
            'Emission Scope 1': scope1_em,
            'Emission Scope 2': scope2_em,
            'Emission Scope 3': scope3_em,
            'Emission Total': total_em,
            'Intensity Scope 1': scope1_int,
            'Intensity Scope 2': scope2_int,
            'Intensity Scope 3': scope3_int,
            'Intensity Total': total_int
        })
        
    df_table2 = pd.DataFrame(results)
    st.subheader("Table 2: Detailed Breakdown")
    st.dataframe(df_table2)

    # ---------------------------------------------------------
    # EXCEL GENERATION
    # ---------------------------------------------------------
    st.write("Generating Excel file with appended dataframes...")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Source Data', index=False)
        df_table1.to_excel(writer, sheet_name='Target Overview', index=False)
        df_table2.to_excel(writer, sheet_name='Target Detailed', index=False)
    
    output.seek(0)
    
    st.download_button(
        label="Download Appended Excel File",
        data=output,
        file_name="GHG_Appended_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
