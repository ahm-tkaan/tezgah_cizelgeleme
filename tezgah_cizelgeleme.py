import pandas as pd
import numpy as np

WORK_POOL_PATH = "veriler/Ürün Havuzu.xlsx"
IMPORTANCE_ORDER_PATH = "veriler/Çizelgeleme Önem Sırası.xlsx"
STANDARD_TIME_PATH = "veriler/Standart süre.xlsx"
INSERT_PATH = "veriler/Ürün kesici uç.xlsx"
MACHINES = ["İM.O01", "İM.O02", "İM.O03", "İM.O04", "İM.O05", "İM.O06", "İM.O07"]
MACHINE_COUNT = len(MACHINES)

def standardize_time_units(standard_time_df):
    df = standard_time_df.copy()
    df["TEZGAHTA ÜRETİM SÜRESİ"] = df["TEZGAHTA ÜRETİM SÜRESİ"].replace(0, 1500).fillna(1500)
    df.loc[df["BİRİM"] == "dakika", "TEZGAHTA ÜRETİM SÜRESİ"] = 1500
    df["SÖK BAĞLA SÜRESİ"] = df["SÖK BAĞLA SÜRESİ"].replace(0, 90).fillna(90)
    df.loc[df["AYAR SÜRESİ BİRİM"] == "dakika", "AYAR SÜRESİ"] *= 60
    df["AYAR SÜRESİ"] = df["AYAR SÜRESİ"].replace(0, 6666).fillna(6666)
    df["BİRİM"] = "saniye"
    df["AYAR SÜRESİ BİRİM"] = "saniye"
    return df

def find_machine(cutting_tip, assignment_rules):
    return assignment_rules.loc[assignment_rules["Kesici Uç"] == cutting_tip]["Tezgah_indeksi"].iloc[0]

def get_machine_indices(cutting_tip, assignment_rules):
    return assignment_rules.loc[assignment_rules["Kesici Uç"] == cutting_tip]["Tezgah_indeksi"].str.split("-").iloc[0]

def get_machine_count(cutting_tip, assignment_rules):
    return assignment_rules.loc[assignment_rules["Kesici Uç"] == cutting_tip]["Kaç Tezgah 2"].iloc[0]

def is_cutting_tip_in_rules(cutting_tip, assignment_rules):
    return cutting_tip in assignment_rules["Kesici Uç"].values

def load_and_prepare_data():
    work_pool = pd.read_excel(WORK_POOL_PATH)
    importance_order = pd.read_excel(IMPORTANCE_ORDER_PATH)
    standard_time = pd.read_excel(STANDARD_TIME_PATH)
    insert_data = pd.read_excel(INSERT_PATH)
    
    standard_time = standardize_time_units(standard_time)
    
    importance_order['Uretim Emri Kodu'] = importance_order['Uretim Emri Kodu'].astype(str)
    work_pool = work_pool.dropna(subset=['IE REFERANS NO'])
    work_pool['IE REFERANS NO'] = work_pool['IE REFERANS NO'].astype(int).astype(str)
    
    insert_data["Stok Adı"] = "DM-" + insert_data["Stok Adı"]
    
    merged_df = pd.merge(work_pool, importance_order, left_on="IE REFERANS NO", right_on="Uretim Emri Kodu", how="left")
    merged_df = pd.merge(merged_df, insert_data, left_on="STOK", right_on="Stok Adı", how="left")
    merged_df = pd.merge(merged_df, standard_time, left_on="STOK", right_on="STOK ADı", how="left")
    
    necessary_cols = ['IE REFERANS NO', 'STOK', 'İE DURUM', 'İŞ EMRİ MİKTARI', 'Sonuc', 
                      "Kesici Uç Kodu 1", "AYAR SÜRESİ", "TEZGAHTA ÜRETİM SÜRESİ", "SÖK BAĞLA SÜRESİ"]
    necessary_data = merged_df[necessary_cols].copy()
    
    necessary_data["Kesici Uç"] = necessary_data["Kesici Uç Kodu 1"].str.split().str[0].str[:4].str.strip()
    necessary_data["İş Süresi"] = (necessary_data["İŞ EMRİ MİKTARI"] * 
                                 (necessary_data["SÖK BAĞLA SÜRESİ"] + necessary_data["TEZGAHTA ÜRETİM SÜRESİ"])) + necessary_data["AYAR SÜRESİ"]
    
    necessary_data = necessary_data.sort_values(by="Sonuc", ascending=False).reset_index(drop=True)
    necessary_data["Tezgah"] = None
    
    return necessary_data

def create_machine_assignment_rules(data, importance_threshold=0.2):
    filtered_data = data[data["Sonuc"] >= importance_threshold]
    
    cutting_tip_totals = filtered_data.groupby('Kesici Uç')['İş Süresi'].sum()
    total_duration = filtered_data['İş Süresi'].sum()
    cutting_tip_percentages = (cutting_tip_totals / total_duration) * 100
    
    percentages_df = cutting_tip_percentages.reset_index()
    percentages_df.columns = ['Kesici Uç', 'Yüzde']
    percentages_df.sort_values(by='Yüzde', ascending=False, inplace=True)
    percentages_df = percentages_df.reset_index(drop=True)
    
    percentages_df["Kaç Tezgah"] = percentages_df["Yüzde"] / (100 / MACHINE_COUNT)
    percentages_df['Kaç Tezgah 2'] = percentages_df['Kaç Tezgah'].apply(
        lambda x: np.ceil(x) if x % 1 > 0.7 else int(x)
    )
    percentages_df['Kaç Tezgah 2'] = percentages_df['Kaç Tezgah 2'].astype(int)
    
    machine_threshold = 0
    machine_index = -1
    machine_assignments = []
    
    for index, row in percentages_df.iterrows():
        remaining_machines = row['Kaç Tezgah 2']
        machine_indices = ""
        machine_threshold += row['Kaç Tezgah']
        
        if remaining_machines == 0:
            if machine_threshold + row['Kaç Tezgah'] > 0.9:
                machine_index += 1
                machine_threshold = 0
            machine_indices = str(machine_index)
        else:
            remaining_machines -= 1
            machine_index += 1
            machine_indices = machine_indices + str(machine_index)
            
            while remaining_machines > 0:
                remaining_machines -= 1
                machine_index += 1
                machine_indices = machine_indices + "-" + str(machine_index)
        
        machine_assignments.append(machine_indices)
    
    while len(machine_assignments) < len(percentages_df):
        machine_assignments.append(machine_assignments[-1])
    
    percentages_df['Tezgah_indeksi'] = machine_assignments
    
    return percentages_df.sort_values(by='Yüzde', ascending=False)

def assign_machines_high_importance(data, assignment_rules):
    assigned_data = data.copy().reset_index(drop=True)
    
    for i in range(len(assigned_data)):
        unique_values = assigned_data["Tezgah"].unique()
        working_machines = unique_values[unique_values != None]
        available_machines = list(set(MACHINES) - set(working_machines))
        
        current_cutting_tip = assigned_data.iloc[i]["Kesici Uç"]
        workload_totals = assigned_data.groupby("Tezgah")["İş Süresi"].sum()
        
        if is_cutting_tip_in_rules(current_cutting_tip, assignment_rules):
            tip_machines = [MACHINES[int(idx)] for idx in get_machine_indices(current_cutting_tip, assignment_rules)]
            common_machines = list(set(available_machines) & set(tip_machines))
            
            if get_machine_count(current_cutting_tip, assignment_rules) > 1:
                if len(common_machines) > 0:
                    assigned_data.iloc[i, assigned_data.columns.get_loc("Tezgah")] = common_machines[0]
                else:
                    workload_totals = workload_totals.reset_index()
                    machine_data = workload_totals[workload_totals['Tezgah'].isin(tip_machines)]
                    
                    if not machine_data.empty:
                        min_workload_machine = machine_data.loc[machine_data["İş Süresi"].idxmin(), "Tezgah"]
                        assigned_data.iloc[i, assigned_data.columns.get_loc("Tezgah")] = min_workload_machine
            else:
                machine_idx = int(find_machine(current_cutting_tip, assignment_rules))
                assigned_data.iloc[i, assigned_data.columns.get_loc("Tezgah")] = MACHINES[machine_idx]
        else:
            if len(available_machines) > 0:
                group_workload = assignment_rules.groupby("Tezgah_indeksi")["Kaç Tezgah"].sum()
                group_workload = group_workload.reset_index()
                min_workload_idx = int(group_workload.loc[group_workload["Kaç Tezgah"].idxmin(), "Tezgah_indeksi"])
                assigned_data.iloc[i, assigned_data.columns.get_loc("Tezgah")] = MACHINES[min_workload_idx]
            else:
                grouped = assigned_data.groupby("Tezgah")
                
                if len(working_machines) > 0:
                    try:
                        group_data = grouped.get_group(working_machines[0])
                        if "Kesici Uç" in group_data.columns and current_cutting_tip == group_data.iloc[-1]["Kesici Uç"]:
                            assigned_data.iloc[i, assigned_data.columns.get_loc("Tezgah")] = working_machines[0]
                        else:
                            order_totals = assigned_data.groupby("Tezgah")["İŞ EMRİ MİKTARI"].sum()
                            min_orders_machine = order_totals.idxmin()
                            assigned_data.iloc[i, assigned_data.columns.get_loc("Tezgah")] = min_orders_machine
                    except:
                        order_totals = assigned_data.groupby("Tezgah")["İŞ EMRİ MİKTARI"].sum()
                        if not order_totals.empty:
                            min_orders_machine = order_totals.idxmin()
                            assigned_data.iloc[i, assigned_data.columns.get_loc("Tezgah")] = min_orders_machine
                        else:
                            assigned_data.iloc[i, assigned_data.columns.get_loc("Tezgah")] = MACHINES[0]
    
    return assigned_data

def assign_machines_low_importance(data, assignment_rules, high_importance_data):
    assigned_data = data.copy().reset_index(drop=True)
    
    machine_utilization = {}
    for machine in MACHINES:
        high_imp_machine_data = high_importance_data[high_importance_data["Tezgah"] == machine]
        machine_utilization[machine] = high_imp_machine_data["İş Süresi"].sum() if not high_imp_machine_data.empty else 0
    
    sorted_machines = sorted(machine_utilization.items(), key=lambda x: x[1])
    
    for i in range(len(assigned_data)):
        current_cutting_tip = assigned_data.iloc[i]["Kesici Uç"]
        
        for machine in MACHINES:
            machine_data = assigned_data[assigned_data["Tezgah"] == machine]
            if not machine_data.empty:
                machine_utilization[machine] = machine_data["İş Süresi"].sum()
        
        sorted_machines = sorted(machine_utilization.items(), key=lambda x: x[1])
        
        if is_cutting_tip_in_rules(current_cutting_tip, assignment_rules):
            high_imp_match = high_importance_data[high_importance_data["Kesici Uç"] == current_cutting_tip]
            
            if not high_imp_match.empty:
                assigned_data.iloc[i, assigned_data.columns.get_loc("Tezgah")] = high_imp_match.iloc[0]["Tezgah"]
            else:
                machine_idx = int(find_machine(current_cutting_tip, assignment_rules))
                if machine_idx < len(MACHINES):
                    assigned_data.iloc[i, assigned_data.columns.get_loc("Tezgah")] = MACHINES[machine_idx]
                else:
                    assigned_data.iloc[i, assigned_data.columns.get_loc("Tezgah")] = sorted_machines[0][0]
        else:
            same_tip_machines = []
            for machine in MACHINES:
                machine_jobs = assigned_data[assigned_data["Tezgah"] == machine]
                if not machine_jobs.empty and current_cutting_tip in machine_jobs["Kesici Uç"].values:
                    same_tip_machines.append(machine)
            
            if same_tip_machines:
                least_utilized = None
                min_util = float('inf')
                for machine in same_tip_machines:
                    if machine_utilization[machine] < min_util:
                        min_util = machine_utilization[machine]
                        least_utilized = machine
                
                assigned_data.iloc[i, assigned_data.columns.get_loc("Tezgah")] = least_utilized
            else:
                assigned_data.iloc[i, assigned_data.columns.get_loc("Tezgah")] = sorted_machines[0][0]
            
            machine = assigned_data.iloc[i]["Tezgah"]
            machine_utilization[machine] += assigned_data.iloc[i]["İş Süresi"]
    
    return assigned_data

def export_to_excel(data, filename="scheduling_results.xlsx"):
    output_data = data.copy()
    
    from datetime import datetime
    output_data["Çizelgeleme Tarihi"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    
    output_columns = {
        'IE REFERANS NO': 'İş Emri Referans No',
        'STOK': 'Stok Kodu',
        'Tezgah': 'Atanan Tezgah',
        'İŞ EMRİ MİKTARI': 'Miktar',
        'Kesici Uç': 'Kesici Uç Kodu',
        'İş Süresi': 'Toplam İş Süresi (sn)',
        'AYAR SÜRESİ': 'Ayar Süresi (sn)',
        'TEZGAHTA ÜRETİM SÜRESİ': 'Üretim Süresi (sn)', 
        'SÖK BAĞLA SÜRESİ': 'Sök-Bağla Süresi (sn)',
        'Sonuc': 'Önem Skoru',
        'Çizelgeleme Tarihi': 'Çizelgeleme Tarihi'
    }
    
    final_data = output_data[list(output_columns.keys())].copy()
    final_data.columns = list(output_columns.values())
    
    machine_summary = final_data.groupby('Atanan Tezgah').agg({
        'İş Emri Referans No': 'count',
        'Toplam İş Süresi (sn)': 'sum',
        'Kesici Uç Kodu': lambda x: len(x.unique())
    }).reset_index()
    
    machine_summary.columns = ['Tezgah', 'İş Emri Sayısı', 'Toplam Süre (sn)', 'Benzersiz Kesici Uç Sayısı']
    machine_summary['Toplam Süre (saat)'] = machine_summary['Toplam Süre (sn)'] / 3600
    
    cutting_tip_summary = final_data.groupby('Kesici Uç Kodu').agg({
        'İş Emri Referans No': 'count',
        'Toplam İş Süresi (sn)': 'sum',
        'Atanan Tezgah': lambda x: ', '.join(sorted(x.unique()))
    }).reset_index()
    
    cutting_tip_summary.columns = ['Kesici Uç', 'İş Sayısı', 'Toplam Süre (sn)', 'Kullanılan Tezgahlar']
    cutting_tip_summary = cutting_tip_summary.sort_values('İş Sayısı', ascending=False)
    
    machine_timeline = final_data.copy()
    machine_timeline = machine_timeline.sort_values(['Atanan Tezgah', 'Önem Skoru'], ascending=[True, False])
    
    machine_timeline['Birikimli Süre (sn)'] = 0
    for machine in machine_timeline['Atanan Tezgah'].unique():
        machine_mask = machine_timeline['Atanan Tezgah'] == machine
        machine_timeline.loc[machine_mask, 'Birikimli Süre (sn)'] = machine_timeline.loc[machine_mask, 'Toplam İş Süresi (sn)'].cumsum()
    
    machine_timeline['Birikimli Süre (saat)'] = machine_timeline['Birikimli Süre (sn)'] / 3600
    
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        final_data.to_excel(writer, sheet_name='Çizelgeleme Sonuçları', index=False)
        machine_summary.to_excel(writer, sheet_name='Tezgah Özeti', index=False)
        cutting_tip_summary.to_excel(writer, sheet_name='Kesici Uç Özeti', index=False)
        machine_timeline.to_excel(writer, sheet_name='Tezgah Zaman Çizelgesi', index=False)
    
    return filename

def main():
    necessary_data = load_and_prepare_data()
    
    high_importance_data = necessary_data[necessary_data["Sonuc"] >= 0.2]
    low_importance_data = necessary_data[necessary_data["Sonuc"] < 0.2]
    
    high_importance_rules = create_machine_assignment_rules(high_importance_data, 0.2)
    high_importance_data_assigned = assign_machines_high_importance(high_importance_data, high_importance_rules)
    
    low_importance_rules = create_machine_assignment_rules(low_importance_data, 0)
    low_importance_data_assigned = assign_machines_low_importance(low_importance_data, low_importance_rules, high_importance_data_assigned)
    
    final_data = pd.concat([high_importance_data_assigned, low_importance_data_assigned])
    
    return final_data

if __name__ == "__main__":
    result = main()
    output_file = export_to_excel(result, "uretim_cizelgeleme_sonuclari.xlsx")
    print(f"Çizelgeleme tamamlandı. {len(result)} iş, {result['Tezgah'].nunique()} tezgaha atandı.")
    print(f"Sonuçlar şu dosyaya kaydedildi: {output_file}")