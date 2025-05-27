import streamlit as st
import pandas as pd
import pulp
from collections import defaultdict
import io

st.set_page_config(page_title="Airline Network Optimizer", layout="centered")
st.title("✈️ Airline Cargo Network Optimizer (Excel Upload)")

uploaded_file = st.file_uploader("📂 Upload your Excel file (.xlsx)", type="xlsx")

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        direct_routes = xls.parse(sheet_name=0).replace("-", 0).fillna(0)
        indirect_routes = xls.parse(sheet_name=1).replace("-", 0).fillna(0)
        st.success("✅ Excel file loaded successfully!")

        all_od_paths = {}
        leg_capacities = {}

        for _, row in direct_routes.iterrows():
            try:
                od = row['O-D']
                cm = float(row['CM'])
                market_share = float(row['AI Share']) if pd.notna(row['AI Share']) else 0
                capacity = float(row['AI Cap']) if pd.notna(row['AI Cap']) else 0
                all_od_paths[od] = {'legs': [od], 'cm': cm, 'max_allocable': min(market_share, capacity)}
                leg_capacities[od] = capacity
            except:
                continue

        for _, row in indirect_routes.iterrows():
            try:
                od = row['O-D']
                cm = float(row['CM'])
                max_allocable = float(row['AI Share']) if pd.notna(row['AI Share']) else 0
                legs = [row['1st Leg O-D'], row['2nd Leg O-D']]
                all_od_paths[od] = {'legs': legs, 'cm': cm, 'max_allocable': max_allocable}
                for i, leg in enumerate(legs):
                    cap_col = f'{i+1}st Leg AI Cap'
                    leg_cap = float(row[cap_col]) if pd.notna(row[cap_col]) else 0
                    leg_capacities[leg] = min(leg_capacities.get(leg, float('inf')), leg_cap)
            except:
                continue

        # Optimization model
        prob = pulp.LpProblem("NetworkCargoProfitMaximization", pulp.LpMaximize)
        x_od = pulp.LpVariable.dicts("CargoTons", all_od_paths.keys(), lowBound=0, cat='Continuous')
        prob += pulp.lpSum([x_od[od] * props['cm'] for od, props in all_od_paths.items()]), "TotalProfit"
        for leg, cap in leg_capacities.items():
            prob += pulp.lpSum([x_od[od] for od, props in all_od_paths.items() if leg in props['legs']]) <= cap
        for od, props in all_od_paths.items():
            prob += x_od[od] <= props['max_allocable']

        prob.solve()

        # OD summary
        od_summary = []
        for v in prob.variables():
            if v.varValue > 0:
                od = v.name.replace("CargoTons_", "").replace("_", "-")
                tons = v.varValue
                cm = all_od_paths[od]['cm']
                profit = tons * cm
                od_summary.append({'OD Pair': od, 'Cargo Tonnage': round(tons, 2), 'CM (₹/ton)': cm, 'Total Profit (₹)': round(profit, 2)})
        df_od_summary = pd.DataFrame(od_summary)

        # Leg breakdown
        records = []
        for od_row in df_od_summary.itertuples():
            od, tons, cm = od_row._1, od_row._2, od_row._3
            for leg in all_od_paths[od]['legs']:
                records.append({
                    'Flight Leg': leg,
                    'OD Contributor': od,
                    'OD CM (₹/ton)': cm,
                    'Cargo Tonnage': tons,
                    'Profit on Leg (₹)': tons * cm
                })

        df_leg_detail = pd.DataFrame(records)

        # Enhanced Priority Classification
        df_leg_detail['Priority Type'] = "Fills Remaining"
        df_leg_detail['Fill Priority Rank'] = None

        for leg, group in df_leg_detail.groupby("Flight Leg"):
            if len(group) == 1:
                idx = group.index[0]
                df_leg_detail.loc[idx, 'Priority Type'] = "Only OD"
                df_leg_detail.loc[idx, 'Fill Priority Rank'] = 1
                continue

            top_cm = group["OD CM (₹/ton)"].max()

            # Sort: Highest CM first, then prefer direct if CM is tied
            sorted_group = group.copy()
            sorted_group['is_direct'] = sorted_group['OD Contributor'].apply(lambda od: len(all_od_paths[od]["legs"]) == 1)
            sorted_group = sorted_group.sort_values(by=["OD CM (₹/ton)", "is_direct"], ascending=[False, False])

            for rank, idx in enumerate(sorted_group.index, start=1):
                row = df_leg_detail.loc[idx]
                od = row["OD Contributor"]
                cm = row["OD CM (₹/ton)"]
                is_direct = len(all_od_paths[od]["legs"]) == 1
                is_highest = cm == top_cm

                if is_highest and is_direct:
                    label = "Highest CM - Direct"
                elif is_highest and not is_direct:
                    label = "Highest CM - Indirect"
                elif is_direct:
                    label = "Direct - Lower CM"
                else:
                    label = "Indirect - Lower CM"

                df_leg_detail.loc[idx, 'Priority Type'] = label
                df_leg_detail.loc[idx, 'Fill Priority Rank'] = rank

        df_leg_summary = df_leg_detail.groupby('Flight Leg')['Cargo Tonnage'].sum().reset_index(name='Total Tonnage (Tons)')
        total_profit = pulp.value(prob.objective)
        df_profit_note = pd.DataFrame([{
            'Flight Leg': 'TOTAL NETWORK PROFIT',
            'OD Contributor': '',
            'OD CM (₹/ton)': '',
            'Cargo Tonnage': '',
            'Profit on Leg (₹)': round(total_profit, 2),
            'Priority Type': '',
            'Fill Priority Rank': ''
        }])

        st.subheader("📦 OD Allocation Summary")
        st.dataframe(df_od_summary)

        st.subheader("✈️ Leg Breakdown (with Priority Type & Rank)")
        st.dataframe(df_leg_detail)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_od_summary.to_excel(writer, index=False, sheet_name="OD_Allocations")
            df_leg_detail.to_excel(writer, index=False, sheet_name="Leg_Breakdown")
            df_leg_summary.to_excel(writer, index=False, sheet_name="Leg_Summary")
            df_profit_note.to_excel(writer, index=False, sheet_name="Profit_Summary")
        output.seek(0)

        st.download_button("📥 Download Excel Report", data=output, file_name="Airline_Cargo_Report.xlsx")

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
