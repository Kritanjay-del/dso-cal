import pandas as pd
df=pd.read_excel("DSO_file.xlsx")
pd.set_option('display.float_format', '{:.2f}'.format)
df.head(16)


import streamlit as st
import pandas as pd


def calculate_dso(row):
    if row['AR'] <= row['Billing']:
        return row['Days'] * row['AR'] / row['Billing']
    
    if pd.isnull(row['AR']):
        return None
    
    total_billing = 0
    net_days = 0
    net_billing = 0
    
    for i in range(row.name, -1, -1):
        total_billing += df.at[i, 'Billing']
        
        if row['AR'] <= total_billing:
            for j in range(i + 1, row.name + 1):
                net_days += df.at[j, 'Days']
                net_billing += df.at[j, 'Billing']
            return round(net_days + (df.at[i, 'Days'] * (row['AR'] - net_billing) / df.at[i, 'Billing']))

    return 'Not Exhaustable'


df['Actual DSO'] = df.apply(calculate_dso, axis=1)


print('*' * 65)
print('*' * 65)


def calculate_billing(row, tar_DSO):
    if tar_DSO <= row['Days']:
        return (row['Billing'] * tar_DSO / row['Days']) - row['Billing']
    
    if tar_DSO <= row['Days'] + df.at[row.name - 1, 'Days']:
        return ((tar_DSO - row['Days']) * row['Billing']) / row['Days']
    
    total_days = 0
    net_days = 0
    net_billing = 0
    
    for i in range(row.name, -1, -1):
        total_days += df.at[i, 'Days']
        
        if tar_DSO <= total_days:
            for j in range(i + 1, row.name + 1):
                net_days += df.at[j, 'Days']
            for k in range(i + 1, row.name):
                net_billing += df.at[k, 'Billing']
            result = round(net_billing + (tar_DSO - net_days) * df.at[i, 'Billing'] / df.at[i, 'Days'])
            
            for col in df.columns:
                print(f"{col}: {row[col]}")
            
            if pd.isnull(result):
                cash_target = None
            elif result == 'Target DSO over 7 months':
                cash_target = 'Target DSO over 7 months'
            elif result <= 0:
                cash_target = "Not Exhaustable"
            else:
                cash_target = df.at[row.name - 1, 'AR'] - result
            
            return f"Cash Target : {round(cash_target)}"


    return 'Target DSO over 7 months'


def calculate_projected_DSO(row, cash_collected):
    if pd.isnull(cash_collected):
        return None

    total_days = 0
    total_billing = 0
    net_billing = 0

    for i in range(row.name - 1, -1, -1):
        total_billing += df.at[i, 'Billing']
        total_days += df.at[i, 'Days']
        
        if (df.at[row.name - 1, 'AR'] - cash_collected) < total_billing:
            for j in range(i + 1, row.name):
                net_billing += df.at[j, 'Billing']
                    

            return round((total_days-df.at[i,'Days'])+df.at[i,'Days'] * (df.at[row.name - 1, 'AR'] - cash_collected - net_billing) / df.at[i, 'Billing'] + df.at[row.name, 'Days'])

    return 'Not Exhaustable'


# Streamlit UI
st.title("DSO Calculator")

month_name = st.text_input("Enter the month name in YYYY-MM format:")
tar_DSO_input = st.number_input("Enter the target DSO:")
cash_collected = st.number_input("Enter the cash collected:")

if st.button("Calculate"):
    filtered_df = df[df['Month'] == month_name]

    if not filtered_df.empty:
        selected_row = filtered_df.iloc[0]
        result = calculate_billing(selected_row, tar_DSO_input)
        projected_dso = calculate_projected_DSO(selected_row, cash_collected)

        try:
            cash_target = float(result.split(': ')[-1])
            percentage_cash_collected = (cash_collected / cash_target) * 100
        except (ValueError, ZeroDivisionError):
            percentage_cash_collected = None

        # Streamlit display with formatting
        # Print 'AR' value
        if isinstance(selected_row['AR'], (int, float)):
            st.write(f"AR for {month_name}: {selected_row['AR']:.0f}")
        else:
            st.write(f"AR for {month_name}: {selected_row['AR']}")

        # Print 'Billing' value
        if isinstance(selected_row['Billing'], (int, float)):
            st.write(f"Billing for {month_name}: {selected_row['Billing']:.0f}")
        else:
            st.write(f"Billing for {month_name}: {selected_row['Billing']}")

        # Print 'Actual DSO' value
        if isinstance(selected_row['Actual DSO'], (int, float)):
            st.write(f"Actual DSO for {month_name}: {selected_row['Actual DSO']:.0f}")
        else:
            st.write(f"Actual DSO for {month_name}: {selected_row['Actual DSO']}")

        # Print 'Cash Target' value
        if 'Cash Target' in result:
            st.write(f"Cash Target: {cash_target:.0f}")

            # Print 'Cash Required' value
            cash_required = cash_target - cash_collected
            if isinstance(cash_required, (int, float)):
                st.write(f"Cash Required to Achieve Target: {cash_required:.0f}")
            else:
                st.write(f"Cash Required to Achieve Target: {cash_required}")
        else:
            st.write("Cannot calculate Cash Required as the Cash Target is not available.")

        # Print 'Projected DSO' value
        if isinstance(projected_dso, (int, float)):
            st.write(f"Projected DSO for {month_name}: {projected_dso:.0f}")
        else:
            st.write(f"Projected DSO for {month_name}: {projected_dso}")

        # Print 'Percentage Collected against Target' value
        st.write(f"Percentage Collected against Target: {percentage_cash_collected:.2f}%")

    else:
        st.write(f"No data found for the month: {month_name}")

