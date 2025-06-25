import pandas as pd
import streamlit as st
from io import BytesIO
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
import plotly.graph_objects as go

# Claim data functions
def filter_claim_data(df):
    return df[df['ClaimStatus'] == 'R']

def remove_duplicate_claims(df):
    dups = df[df.duplicated(subset='ClaimNo', keep=False)]
    if not dups.empty:
        st.write("Duplicated ClaimNo values:")
        st.dataframe(dups[['ClaimNo']].drop_duplicates())
    return df.drop_duplicates(subset='ClaimNo', keep='last')

def process_claim_data(df):
    df = filter_claim_data(df)
    df = remove_duplicate_claims(df)
    for col in ["TreatmentStart", "TreatmentFinish", "Date"]:
        df[col] = pd.to_datetime(df[col], errors='coerce')
        if df[col].isnull().any():
            st.warning(f"Invalid date values detected in column '{col}'. Coerced to NaT.")

    # Build standardized template:
    return pd.DataFrame({
        "No": range(1, len(df) + 1),
        "Policy No": df["PolicyNo"],
        "Client Name": df["ClientName"],
        "Claim No": df["ClaimNo"],
        "Member No": df["MemberNo"],
        "Emp ID": df["EmpID"],
        "Emp Name": df["EmpName"],
        "Patient Name": df["PatientName"],
        "Membership": df["Membership"],
        "Product Type": df["ProductType"],
        "Claim Type": df["ClaimType"],
        "Room Option": df["RoomOption"].fillna('').astype(str).str.upper().str.replace(r"\s+", "", regex=True),
        "Area": df["Area"],
        "Plan": df["PPlan"],
        "Diagnosis": df["PrimaryDiagnosis"].str.upper(),
        "Treatment Place": df["TreatmentPlace"].str.upper(),
        "Treatment Start": df["TreatmentStart"],
        "Treatment Finish": df["TreatmentFinish"],
        "Settled Date": df["Date"],
        "Year": df["Date"].dt.year,
        "Month": df["Date"].dt.month,
        "Length of Stay": df["LOS"],
        "Sum of Billed": df["Billed"],
        "Sum of Accepted": df["Accepted"],
        "Sum of Excess Coy": df["ExcessCoy"],
        "Sum of Excess Emp": df["ExcessEmp"],
        "Sum of Excess Total": df["ExcessTotal"],
        "Sum of Unpaid": df["Unpaid"]
    })

# Benefit data functions
def filter_benefit_data(df):
    if 'Status_Claim' in df.columns:
        return df[df['Status_Claim'] == 'R']
    elif 'Status Claim' in df.columns:
        return df[df['Status Claim'] == 'R']
    else:
        st.warning("Column 'Status Claim' not found. Data not filtered.")
        return df

def process_benefit_data(df):
    df = filter_benefit_data(df)
    df.columns = df.columns.str.strip()
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].astype(str).str.strip()
    rename_mapping = {
        'ClientName': 'Client Name',
        'PolicyNo': 'Policy No',
        'ClaimNo': 'Claim No',
        'MemberNo': 'Member No',
        'PatientName': 'Patient Name',
        'EmpID': 'Emp ID',
        'EmpName': 'Emp Name',
        'ClaimType': 'Claim Type',
        'TreatmentPlace': 'Treatment Place',
        'RoomOption': 'Room Option',
        'TreatmentRoomClass': 'Treatment Room Class',
        'TreatmentStart': 'Treatment Start',
        'TreatmentFinish': 'Treatment Finish',
        'ProductType': 'Product Type',
        'BenefitName': 'Benefit Name',
        'PaymentDate': 'Payment Date',
        'ExcessTotal': 'Excess Total',
        'ExcessCoy': 'Excess Coy',
        'ExcessEmp': 'Excess Emp'
    }
    df = df.rename(columns=rename_mapping)
    return df.drop(columns=["Status_Claim", "BAmount"], errors='ignore')

# Save to excel
def save_to_excel(claim_df, benefit_df, summary_top_df, claim_ratio_df, filename):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        workbook.formats[0].set_font_name('VAG Rounded Std Light')

        # Define excel formats:
        bold_border = workbook.add_format({'bold': True, 'border': 1, 'font_name': 'VAG Rounded Std Light'})
        plain_border = workbook.add_format({'border': 1, 'font_name': 'VAG Rounded Std Light'})
        header_border = workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'font_name': 'VAG Rounded Std Light'})

        # Summary sheet:
        summary_sheet = workbook.add_worksheet("Summary")
        summary_sheet.hide_gridlines(2)
        row = 0
        # Write summary statistics:
        for _, data in summary_top_df.iterrows():
            summary_sheet.write(row, 0, data["Metric"], bold_border)
            summary_sheet.write(row, 1, data["Value"], plain_border)
            row += 1
        # Insert blank row without borders (seperate sum stats & CR)
        summary_sheet.write(row, 0, "")
        summary_sheet.write(row, 1, "")
        row += 1

        # Write header for Claim Ratio table
        cr_columns = ["Company", "Net Premi", "Billed", "Unpaid", "Excess Total", "Excess Coy", "Excess Emp", "Claim", "CR", "Est Claim"]
        for col, header in enumerate(cr_columns):
            summary_sheet.write(row, col, header, header_border)
        row += 1

        # Write Claim Ratio data rows
        for _, data in claim_ratio_df.iterrows():
            for col, header in enumerate(cr_columns):
                summary_sheet.write(row, col, data.get(header, ""), plain_border)
            row += 1

        # Claim (SC) sheet
        claim_df.to_excel(writer, index=False, sheet_name='SC')
        ws_claim = writer.sheets["SC"]
        ws_claim.hide_gridlines(2)
        rows_claim, cols_claim = claim_df.shape[0] + 1, claim_df.shape[1]
        ws_claim.conditional_format(0, 0, rows_claim - 1, cols_claim - 1,
                                     {'type': 'no_errors', 'format': plain_border})
        for col_num, value in enumerate(claim_df.columns.values):
            ws_claim.write(0, col_num, value, header_border)

        # Benefit sheet
        benefit_df.to_excel(writer, index=False, sheet_name='Benefit')
        ws_benefit = writer.sheets["Benefit"]
        ws_benefit.hide_gridlines(2)
        rows_benefit, cols_benefit = benefit_df.shape[0] + 1, benefit_df.shape[1]
        ws_benefit.conditional_format(0, 0, rows_benefit - 1, cols_benefit - 1,
                                      {'type': 'no_errors', 'format': plain_border})
        for col_num, value in enumerate(benefit_df.columns.values):
            ws_benefit.write(0, col_num, value, header_border)


        writer.close()
    output.seek(0)
    return output, filename

# Streamlit APP UI
st.title("Template - Standardisasi Report")

uploaded_claim = st.file_uploader("Upload Claim Data", type=["csv"], key="claim")
uploaded_claim_ratio = st.file_uploader("Upload Claim Ratio Data", type=["xlsx"], key="claim_ratio")
uploaded_benefit = st.file_uploader("Upload Benefit Data", type=["csv"], key="benefit")

if uploaded_claim and uploaded_claim_ratio and uploaded_benefit:
    # Process claim data
    raw_claim = pd.read_csv(uploaded_claim)
    st.write("Processing Claim Data...")
    claim_transformed = process_claim_data(raw_claim)
    st.subheader("Claim Data Preview:")
    st.dataframe(claim_transformed.head())

    # Process claim ratio data
    claim_ratio_raw = pd.read_excel(uploaded_claim_ratio)
    policy_list = claim_transformed["Policy No"].unique().tolist()
    claim_ratio_filtered = claim_ratio_raw[claim_ratio_raw["Policy No"].isin(policy_list)]
    claim_ratio_unique = claim_ratio_filtered.drop_duplicates(subset="Policy No")
    desired_cols = ['Company', 'Net Premi', 'Billed', 'Unpaid',
                    'Excess Total', 'Excess Coy', 'Excess Emp', 'Claim', 'CR', 'Est CR Total']
    missing_cols = [col for col in desired_cols if col not in claim_ratio_unique.columns]
    if missing_cols:
        st.warning(f"Missing columns in Claim Ratio Data: {missing_cols}")
    claim_ratio_unique = claim_ratio_unique[[col for col in desired_cols if col in claim_ratio_unique.columns]]
    claim_ratio_unique = claim_ratio_unique.rename(columns={'Est CR Total': 'Est Claim'})
    summary_cr_df = claim_ratio_unique[['Company', 'Net Premi', 'Billed', 'Unpaid',
                                         'Excess Total', 'Excess Coy', 'Excess Emp', 'Claim', 'CR', 'Est Claim']]
    st.subheader("Claim Ratio Data Preview (unique by Policy No):")
    st.dataframe(summary_cr_df.head())

    # Process benefit data
    raw_benefit = pd.read_csv(uploaded_benefit)
    st.write("Processing Benefit Data...")
    benefit_transformed = process_benefit_data(raw_benefit)
    claim_no_list = claim_transformed["Claim No"].unique().tolist()
    if "ClaimNo" in benefit_transformed.columns:
        benefit_transformed = benefit_transformed[benefit_transformed["ClaimNo"].isin(claim_no_list)]
    elif "Claim No" in benefit_transformed.columns:
        benefit_transformed = benefit_transformed[benefit_transformed["Claim No"].isin(claim_no_list)]
    else:
        st.warning("Column 'ClaimNo' not found in Benefit data; skipping filtering based on ClaimNo.")
    st.subheader("Benefit Data Preview:")
    st.dataframe(benefit_transformed.head())

    # Prepare Summary Top Section (Claim Stats + Overall Claim Ratio)
    total_claims   = len(claim_transformed)
    total_billed   = int(claim_transformed["Sum of Billed"].sum())
    total_accepted = int(claim_transformed["Sum of Accepted"].sum())
    total_excess   = int(claim_transformed["Sum of Excess Total"].sum())
    total_unpaid   = int(claim_transformed["Sum of Unpaid"].sum())
    claim_summary_data = {
        "Metric": ["Total Claims", "Total Billed", "Total Accepted", "Total Excess", "Total Unpaid"],
        "Value": [f"{total_claims:,}", f"{total_billed:,}", f"{total_accepted:,}",
                  f"{total_excess:,}", f"{total_unpaid:,}"]
    }
    claim_summary_df = pd.DataFrame(claim_summary_data)

    if "Claim" in claim_ratio_unique.columns and "Net Premi" in claim_ratio_unique.columns:
        total_claim_ratio_claim = claim_ratio_unique["Claim"].sum()
        total_net_premi = claim_ratio_unique["Net Premi"].sum()
        overall_cr = (total_claim_ratio_claim / total_net_premi) * 100 if total_net_premi != 0 else 0
        claim_ratio_overall = pd.DataFrame({"Metric": ["Claim Ratio (%)"],
                                            "Value": [f"{overall_cr:.2f}%"]})
    else:
        claim_ratio_overall = pd.DataFrame({"Metric": ["Claim Ratio (%)"], "Value": ["N/A"]})

    summary_top_df = pd.concat([claim_summary_df, claim_ratio_overall], ignore_index=True)

    # Download the Excel file
    st.subheader("Download Processed Data")
    filename_input = st.text_input("Enter the Excel file name (without extension):", "SC & Benefit - - YTD")
    if filename_input:
        excel_file, final_filename = save_to_excel(claim_transformed, benefit_transformed,
                                                   summary_top_df, summary_cr_df, filename_input + ".xlsx")
        st.download_button(
            label="Download Processed Excel File",
            data=excel_file,
            file_name=final_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # Display Visualizations Section (Moved after download button)
    st.subheader("Visualizations")

    # Section 1: Summary Metrics Visualization
    st.subheader("Summary Metrics")
    # Prepare data for horizontal view
    metrics_row1 = summary_top_df.iloc[0:3]
    metrics_row2 = summary_top_df.iloc[3:]

    # Function to display metric with formatted text
    def display_metric(metric_name, metric_value):
        # Coba ubah ke float lalu ke integer jika bisa, agar menghapus desimal
        try:
            formatted_value = f"{int(float(str(metric_value).replace(',', ''))):,}"
        except:
            formatted_value = metric_value  # fallback ke original
        st.markdown(f"<p style='color: #0067B1; font-size: 18px; margin-bottom: 0;'>{metric_name}</p>"
                    f"<p style='color: #0067B1; font-size: 24px; font-weight: bold; margin-top: 0;'>{metric_value}</p>",
                    unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)
    with col1:
        display_metric(metrics_row1.iloc[0]['Metric'], metrics_row1.iloc[0]['Value'])
    with col2:
        display_metric(metrics_row1.iloc[1]['Metric'], metrics_row1.iloc[1]['Value'])
    with col3:
        display_metric(metrics_row1.iloc[2]['Metric'], metrics_row1.iloc[2]['Value'])

    col4, col5, col6 = st.columns(3)
    with col4:
        display_metric(metrics_row2.iloc[0]['Metric'], metrics_row2.iloc[0]['Value'])
    with col5:
        display_metric(metrics_row2.iloc[1]['Metric'], metrics_row2.iloc[1]['Value'])
    with col6:
        display_metric(metrics_row2.iloc[2]['Metric'], metrics_row2.iloc[2]['Value']) # Assuming 6 metrics in total, adjust if needed

    # Claim Ratio Summary Table (Using HTML/CSS for enhanced display)
    st.subheader("Claim Ratio Summary Table")

    def format_claim_ratio_table(df):
        html = "<style>"
        html += "table { border-collapse: collapse; width: 100%; }"
        html += "th, td { border: 1px solid #333; padding: 8px; text-align: center; color: black; }"
        html += "th { background-color: #0070C0; font-weight: bold; color: white; }"
        html += "tr:nth-child(even) { background-color: #fcfcfa; }"
        html += "tr:hover { background-color: #ddd; }"
        html += "</style>"
        html += "<table><thead><tr>"
    
        # Table headers
        for col in df.columns:
            html += f"<th>{col}</th>"
        html += "</tr></thead><tbody>"
    
        # Table rows
        for _, row in df.iterrows():
            html += "<tr>"
            for col in df.columns:
                value = row[col]
                if isinstance(value, (int, float)):
                    if col == 'CR' and not pd.isna(value):
                        content = f"{value:.2f}%"
                    elif col == 'Est Claim' and not pd.isna(value):
                        content = f"{value:,.2f}"
                    else:
                        content = f"{int(value):,}"  # Pastikan bulat tanpa desimal
                else:
                    content = value
                html += f"<td>{content}</td>"
            html += "</tr>"
        html += "</tbody></table>"
        return html



    st.markdown(format_claim_ratio_table(summary_cr_df), unsafe_allow_html=True)
    

    # Section 2: Claim per Membership (Pie Chart)
    st.subheader("Claim Count per Membership Type")
    
    # Cek apakah kolom 'Membership' tersedia
    if 'Membership' in claim_transformed.columns:
        # Mapping label Membership
        label_map = {
            "1. EMP": "Employee",
            "2. SPO": "Spouse",
            "3. CHI": "Children"
        }
        claim_transformed['Membership Label'] = claim_transformed['Membership'].map(label_map)
    
        # Hitung jumlah klaim per jenis Membership
        membership_counts = claim_transformed['Membership Label'].value_counts().reset_index()
        membership_counts.columns = ['Membership Type', 'Claim Count']
        membership_counts = membership_counts.sort_values(by='Membership Type')
    
        labels = membership_counts['Membership Type'].tolist()
        values = membership_counts['Claim Count'].tolist()
        total = sum(values)
        text_labels = [f"<b>{(v/total)*100:.0f}%</b><br>({v} claim)" for v in values]
    
        # Warna sesuai urutan biru (dari gelap ke terang)
        colors = ['#1f77b4', '#4e91c7', '#a6c8ea']
    
        # Membuat pie chart
        fig = go.Figure(data=[go.Pie(
            labels=labels,
            values=values,
            marker=dict(colors=colors),
            text=text_labels,
            textinfo='text',
            insidetextorientation='horizontal',
            hoverinfo='label+value',
            pull=[0, 0, 0],
            showlegend=True,
            textfont=dict(color='white', size=14, family='Arial'),
            hole=0
        )])
    
        # Layout chart
        fig.update_layout(
            legend_orientation="h",
            legend_title_text='',
            legend=dict(
                yanchor="top",
                y=-0.1,
                xanchor="center",
                x=0.5
            ),
            margin=dict(t=60, b=60),
            height=400,
            width=400
        )
    
        st.plotly_chart(fig)
    else:
        st.warning("'Membership' column not found in Claim Data.")
    
    # Section 3: Claim Count per Plan
    st.subheader("Claim Count per Plan")
    
    # Cek apakah kolom 'Plan' tersedia
    if 'Plan' in claim_transformed.columns:
        # Hitung jumlah klaim per jenis Plan
        plan_counts = claim_transformed['Plan'].value_counts().reset_index()
        plan_counts.columns = ['Plan', 'Claim Count']
        plan_counts = plan_counts.sort_values(by='Plan')
    
        # Membuat bar chart
        fig = go.Figure()
    
        # Tambahkan bar chart
        fig.add_trace(go.Bar(
            x=plan_counts['Plan'],
            y=plan_counts['Claim Count'],
            marker_color='#1f77b4'  # Warna biru
        ))
    
        # Tambahkan teks jumlah di atas bar
        for i in range(len(plan_counts)):
            fig.add_annotation(
                x=plan_counts['Plan'][i],
                y=plan_counts['Claim Count'][i],
                text=str(plan_counts['Claim Count'][i]),
                showarrow=False,
                yshift=10,  # Posisikan teks sedikit di atas bar
                font=dict(
                    color="black",  # Warna teks hitam
                    size=12
                )
            )
    
        # Layout chart
        fig.update_layout(
            xaxis_title="Plan",
            yaxis_title="Number of Claims",
            margin=dict(t=60, b=60),
            height=400,
            width=600,
            font=dict(color='black'),
            xaxis=dict(
                title_font=dict(color='black'),
                tickfont=dict(color='black')
            ),
            yaxis=dict(
                title_font=dict(color='black'),
                tickfont=dict(color='black')
            )
        )

        st.plotly_chart(fig)
    
    else:
        st.warning("'Plan' column not found in Claim Data.")

    # Section 4: Claim Billed by Month and Product Type
    st.subheader("Claim Billed by Month and Product Type")
    
    # Ensure 'Settled Date' and 'Product Type' columns exist
    if 'Settled Date' in claim_transformed.columns and 'Product Type' in claim_transformed.columns:
        # Convert 'Settled Date' to datetime and create month-year column
        claim_transformed['Settled Month'] = claim_transformed['Settled Date'].dt.strftime("%b'%y")
    
        # Group by month and product type and sum the billed amount
        monthly_billed_by_product = claim_transformed.groupby(['Settled Month', 'Product Type'])['Sum of Billed'].sum().reset_index()
    
        # Get the correct month order
        month_order = pd.to_datetime(claim_transformed['Settled Date']).dt.to_period('M').sort_values().unique().strftime("%b'%y").tolist()
    
        # Convert 'Settled Month' to categorical with the defined order for correct sorting
        monthly_billed_by_product['Settled Month'] = pd.Categorical(monthly_billed_by_product['Settled Month'], categories=month_order, ordered=True)
    
        # Sort by month
        monthly_billed_by_product = monthly_billed_by_product.sort_values('Settled Month')
    
        # Create the grouped bar chart using Plotly
        fig = px.bar(monthly_billed_by_product,
                     x='Settled Month',
                     y='Sum of Billed',
                     color='Product Type',
                     barmode='group',  # Use 'group' for grouped bars
                     labels={'Settled Month': 'Settled Month', 'Sum of Billed': 'Sum of Billed Amount'},
                     title='Claim Billed by Month and Product Type')
    
        # Update layout for better appearance
        fig.update_layout(
            xaxis_title="Settled Month",
            yaxis_title="Sum of Billed Amount",
            margin=dict(t=60, b=10), # Reduce bottom margin to bring chart closer to the table
            font=dict(color='black'),
            xaxis=dict(
                title_font=dict(color='black'),
                tickfont=dict(color='black')
            ),
            yaxis=dict(
                title_font=dict(color='black'),
                tickfont=dict(color='black')
            ),
            legend_title_text='Product Type',
            height=400 # Adjust height if needed
        )
    
        # Display the chart
        st.plotly_chart(fig, use_container_width=True) # use_container_width=True makes it responsive
    
        # Display the detailed table below the chart
        st.subheader("Claim Billed Details by Month and Product Type")
        # Pivot the table for better readability if needed, or display as is
        # Pivoting to show months as index and product types as columns
        pivot_table = monthly_billed_by_product.pivot_table(index='Settled Month', columns='Product Type', values='Sum of Billed', fill_value=0).reset_index()
        pivot_table.columns.name = None # Remove the columns name 'Product Type'
    
        # Re-sort the pivoted table by the ordered month index
        pivot_table['Settled Month'] = pd.Categorical(pivot_table['Settled Month'], categories=month_order, ordered=True)
        pivot_table = pivot_table.sort_values('Settled Month')
    
        st.dataframe(pivot_table)
    
    else:
        st.warning("'Settled Date' or 'Product Type' column not found in Claim Data. Cannot generate Section 4 visualization.")

    # Section 5: Top 10 Diagnoses by Product Type
    st.subheader("Top 10 Diagnoses by Product Type")
    
    # Warna
    color_amount = '#1f77b4'  # Dark blue
    color_qty = '#a6c8ea'     # Light blue
    
    # Grouping
    diagnosis_summary = claim_transformed.groupby(['Product Type', 'Diagnosis']).agg(
        Amount=('Sum of Billed', 'sum'),
        Qty=('Sum of Billed', 'count')
    ).reset_index()
    
    # Scale to millions
    diagnosis_summary['Amount'] = diagnosis_summary['Amount'] / 1_000_000
    
    # Loop per product type
    for product in diagnosis_summary['Product Type'].unique():
        st.markdown(f"### {product}")
    
        top_10 = (
            diagnosis_summary[diagnosis_summary['Product Type'] == product]
            .sort_values(by='Amount', ascending=False)
            .head(10)
        )
    
        fig = go.Figure()
    
        # Trace 1: Qty
        fig.add_trace(go.Bar(
            y=top_10['Diagnosis'],
            x=top_10['Qty'],
            name='Qty',
            orientation='h',
            marker_color=color_qty,
            text=[f"{v:,}" for v in top_10['Qty']],
            textposition='outside',
            textfont=dict(color='black'),
            legendgroup='qty',
            legendrank=2
        ))
    
        # Trace 2: Amount
        fig.add_trace(go.Bar(
            y=top_10['Diagnosis'],
            x=top_10['Amount'],
            name='Amount (in millions)',
            orientation='h',
            marker_color=color_amount,
            text=[f"{v:,.2f}" if v < 1 else f"{v:,.0f}" for v in top_10['Amount']],
            textposition='outside',
            textfont=dict(color='black'),
            legendgroup='amount',
            legendrank=1
        ))
    
        # Layout
        fig.update_layout(
            barmode='group',
            yaxis=dict(
                categoryorder='total ascending',
                title='Diagnosis',
                title_font=dict(color='black'),
                tickfont=dict(color='black')
            ),
            xaxis=dict(
                title='Value',
                title_font=dict(color='black'),
                tickfont=dict(color='black')
            ),
            font=dict(color='black'),  # General text color (legend, etc.)
            legend_title_text='',
            height=400,
            margin=dict(t=40, b=40),
            bargap=0.2
        )
    
        st.plotly_chart(fig, use_container_width=True)

    # Section 6: Top 10 Treatment Places by Claim Type
    st.subheader("Top 10 Treatment Places by Claim Type")

    # Warna tetap sama
    color_amount = '#1f77b4'  # Dark blue
    color_qty = '#a6c8ea'     # Light blue
    
    # Grouping
    treatment_place_summary = claim_transformed.groupby(['Claim Type', 'Treatment Place']).agg(
        Amount=('Sum of Billed', 'sum'),
        Qty=('Sum of Billed', 'count')
    ).reset_index()
    
    # Scale to millions
    treatment_place_summary['Amount'] = treatment_place_summary['Amount'] / 1_000_000
    
    # Loop per Claim Type
    for claim_type in treatment_place_summary['Claim Type'].unique():
        st.markdown(f"### {claim_type}")
    
        top_10 = (
            treatment_place_summary[treatment_place_summary['Claim Type'] == claim_type]
            .sort_values(by='Amount', ascending=False)
            .head(10)
        )
    
        fig = go.Figure()
    
        # Trace 1: Qty (first → appears below)
        fig.add_trace(go.Bar(
            y=top_10['Treatment Place'],
            x=top_10['Qty'],
            name='Qty',
            orientation='h',
            marker_color=color_qty,
            text=[f"{v:,}" for v in top_10['Qty']],
            textposition='outside',
            textfont=dict(color='black'),
            legendgroup='qty',
            legendrank=2
        ))
    
        # Trace 2: Amount (second → appears above)
        fig.add_trace(go.Bar(
            y=top_10['Treatment Place'],
            x=top_10['Amount'],
            name='Amount (in millions)',
            orientation='h',
            marker_color=color_amount,
            text=[f"{v:,.2f}" if v < 1 else f"{v:,.0f}" for v in top_10['Amount']],
            textposition='outside',
            textfont=dict(color='black'),
            legendgroup='amount',
            legendrank=1
        ))
    
        # Layout
        fig.update_layout(
            barmode='group',
            yaxis=dict(
                categoryorder='total ascending',
                title='Treatment Place',
                title_font=dict(color='black'),
                tickfont=dict(color='black')
            ),
            xaxis=dict(
                title='Value',
                title_font=dict(color='black'),
                tickfont=dict(color='black')
            ),
            font=dict(color='black'),
            legend_title_text='',
            height=400,
            margin=dict(t=40, b=40),
            bargap=0.2
        )
    
    st.plotly_chart(fig, use_container_width=True)

    # Section 7: Top 10 Employee
    st.subheader("Top 10 Employees by Number of Claims")

    df_emp = claim_transformed.copy()
    
    # Group and summarize
    top_10_emp_summary = (
        df_emp.groupby(['Emp Name', 'Plan'])
              .agg(
                  Total_Claims=('Emp Name', 'count'),
                  Total_Billed=('Sum of Billed', 'sum')
              )
              .reset_index()
    )
    
    # Show warning if empty
    if top_10_emp_summary.empty:
        st.warning("No employee data available.")
    else:
        # Sort & take top 10
        top_10_emp_summary = top_10_emp_summary.sort_values(by='Total_Claims', ascending=False).head(10)
    
        # Rename and reorder
        top_10_emp_summary = top_10_emp_summary.rename(columns={
            'Emp Name': 'Employee',
            'Total_Claims': 'Total Claims',
            'Total_Billed': 'Total Billed'
        })[['Employee', 'Plan', 'Total Claims', 'Total Billed']]
    
        # Format numbers
        top_10_emp_summary['Total Claims'] = top_10_emp_summary['Total Claims'].map('{:,}'.format)
        top_10_emp_summary['Total Billed'] = top_10_emp_summary['Total Billed'].map('{:,}'.format)
    
        # Styled HTML table with thick visible borders
        def render_styled_table(df):
            html = """
            <style>
            table {
                border-collapse: collapse;
                width: 100%;
                font-family: Arial, sans-serif;
            }
            tr:nth-child(even) {
                background-color: #f5f5f5;
            }
            </style>
            <table>
                <thead>
                    <tr>
            """
            # Tambahkan header dengan inline border
            for col in df.columns:
                html += f"<th style='border: 1px solid #333; background-color: #0067B1; color: white; padding: 8px;'>{col}</th>"
            html += "</tr></thead><tbody>"
        
            # Tambahkan isi tabel dengan border per cell
            for _, row in df.iterrows():
                html += "<tr>"
                for item in row:
                    html += f"<td style='border: 1px solid #333; color: black; padding: 8px; text-align: center;'>{item}</td>"
                html += "</tr>"
        
            html += "</tbody></table>"
            return html
    
        # Render in Streamlit
        st.markdown(render_styled_table(top_10_emp_summary), unsafe_allow_html=True)
