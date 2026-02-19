# import pandas as pd
# import json
# import re
# import sys
# import os

# # ─── Load JSON reference data ───────────────────────────────────────────────
# def load_json(path):
#     with open(path) as f:
#         return json.load(f)

# BASE = r'C:\Users\suyash.kamath\Desktop\Liberty-Payin-Configuration'
# companies   = load_json(f'{BASE}/company_master.json')
# segments    = load_json(f'{BASE}/segment.json')
# subproducts = load_json(f'{BASE}/subproduct.json')
# fuel_types  = load_json(f'{BASE}/fuel.json')
# vehicle_types = load_json(f'{BASE}/vehicle_type.json')
# rto_list    = load_json(f'{BASE}/rto_id_name.json')

# # Build lookup dicts
# company_dict    = {c['company_id']: c for c in companies}
# segment_dict    = {'Comprehensive': (1, 'Comprehensive'), 'SAOD': (2, 'SAOD'), 'TP Only': (3, 'TP Only')}
# policy_map      = {'COMP': 'Comprehensive', 'TP': 'TP Only', 'SAOD': 'SAOD'}
# subprod_dict    = {s['sub_product_name']: s['sub_product_id'] for s in subproducts}
# # LOB -> sub_product_name mapping (from the input LOB column)
# lob_map         = {'TW': 'Two Wheeler', 'PC': 'Private Car', 'GCV': 'Goods Vehicle', 'PCV': 'Passenger Vehicle'}
# # vehicle type lookup: (lob, tw_type) -> vehicle_type_id
# vt_dict         = {(vt['sub_product_name'], vt['vehicle_type']): vt['id'] for vt in vehicle_types}
# # RTO lookup
# rto_dict        = {r['name']: r['id'] for r in rto_list}

# # ─── Company selection ───────────────────────────────────────────────────────
# print("\n" + "="*80)
# print("AVAILABLE COMPANIES")
# print("="*80)
# for c in companies:
#     print(f"  ID: {c['company_id']:3d} | Code: {c['company_code']:20s} | {c['company_name']}")
# print("="*80)

# company_id_input = input("\nEnter company_id from the list above: ").strip()
# try:
#     company_id = int(company_id_input)
#     company = company_dict[company_id]
#     print(f"\n✓ Selected: {company['company_name']} (ID: {company_id}, Code: {company['company_code']})")
# except (ValueError, KeyError):
#     print(f"ERROR: Invalid company_id '{company_id_input}'")
#     sys.exit(1)

# # ─── CC Band parser ──────────────────────────────────────────────────────────
# def parse_cc_band(cc_band):
#     """Returns (from_cc, to_cc, is_cc_considered)"""
#     if not cc_band or str(cc_band).strip() == '' or str(cc_band).strip().lower() == 'nan':
#         return 0, 99999, -1
    
#     s = str(cc_band).strip().upper().replace('CC', '').strip()
    
#     # Pattern: < 75 -> 0 to 74
#     m = re.match(r'^<\s*(\d+)$', s)
#     if m:
#         val = int(m.group(1))
#         return 0, val - 1, 1
    
#     # Pattern: > 350 (no range) -> 351 to 99999
#     m = re.match(r'^>\s*(\d+)$', s)
#     if m:
#         val = int(m.group(1))
#         return val + 1, 99999, 1
    
#     # Pattern: > 150-350 -> 151 to 350
#     m = re.match(r'^>\s*(\d+)\s*[-–]\s*(\d+)$', s)
#     if m:
#         lo, hi = int(m.group(1)), int(m.group(2))
#         return lo + 1, hi, 1
    
#     # Pattern: >= 150-350 -> 150 to 350
#     m = re.match(r'^>=\s*(\d+)\s*[-–]\s*(\d+)$', s)
#     if m:
#         lo, hi = int(m.group(1)), int(m.group(2))
#         return lo, hi, 1
    
#     # Pattern: 75-150 -> 75 to 150
#     m = re.match(r'^(\d+)\s*[-–]\s*(\d+)$', s)
#     if m:
#         lo, hi = int(m.group(1)), int(m.group(2))
#         return lo, hi, 1
    
#     # Pattern: single number -> from that to that
#     m = re.match(r'^(\d+)$', s)
#     if m:
#         val = int(m.group(1))
#         return val, val, 1
    
#     # Fallback
#     return 0, 99999, -1

# # ─── Vehicle type ID resolver ────────────────────────────────────────────────
# def get_vehicle_type_id(lob, tw_type):
#     """Map LOB + TW Type to vehicle_type_id"""
#     sub_prod = lob_map.get(lob, lob)
    
#     # TW specific
#     if sub_prod == 'Two Wheeler':
#         if tw_type and 'bike' in str(tw_type).lower():
#             # Check electric or regular
#             return 18  # TW Bike
#         elif tw_type and 'scooter' in str(tw_type).lower():
#             return 17  # TW Scooter
#     return -1

# # ─── is_cc_considered from Original Segment ──────────────────────────────────
# def get_is_cc_considered(original_segment):
#     """Check if numbers appear in the original segment string"""
#     if not original_segment or str(original_segment).strip().lower() == 'nan':
#         return -1
#     # If there's any digit in the string, cc is mentioned
#     if re.search(r'\d', str(original_segment)):
#         return 1
#     return -1

# # ─── Main processing ─────────────────────────────────────────────────────────
# df = pd.read_excel("C:/Users/suyash.kamath/Downloads/TW_Processed_TW Comp_R_1_1771413426366.xlsx")
# print(f"\nProcessing {len(df)} records from input file...")

# rows = []
# for idx, row in df.iterrows():
#     # Company
#     comp_id   = company_id
#     comp_code = company['company_code']
    
#     # Segment
#     policy_type = str(row.get('Policy Type', '')).strip()
#     seg_name_raw = policy_map.get(policy_type, 'TP Only')
#     seg_id, seg_name = segment_dict.get(seg_name_raw, (3, 'TP Only'))
    
#     # Subproduct
#     lob = str(row.get('LOB', 'TW')).strip()
#     sub_prod_name = lob_map.get(lob, lob)
#     sub_prod_id   = subprod_dict.get(sub_prod_name, -1)
    
#     # RTO
#     geo_new = str(row.get('Geo New', '')).strip()
#     rto_group_id   = rto_dict.get(geo_new, -1)
#     rto_group_name = geo_new
    
#     # Payin/Payout rates
#     payin_val = row.get('Payin', 0)
#     payout_val = row.get('Calculated Payout', 0)
    
#     # Convert percentage strings to float if needed
#     def to_float(v):
#         if isinstance(v, str):
#             v = v.strip().replace('%', '')
#             try:
#                 return float(v) / 100
#             except:
#                 return 0.0
#         try:
#             return float(v)
#         except:
#             return 0.0
    
#     payin_val  = to_float(payin_val)
#     payout_val = to_float(payout_val)
    
#     if policy_type == 'TP':
#         payin_od_rate   = 0
#         payin_tp_rate   = payin_val
#         payout_od_rate  = 0
#         payout_tp_rate  = payout_val
#     elif policy_type == 'SAOD':
#         payin_od_rate   = payin_val
#         payin_tp_rate   = 0
#         payout_od_rate  = payout_val
#         payout_tp_rate  = 0
#     else:  # COMP
#         payin_od_rate   = payin_val
#         payin_tp_rate   = payin_val
#         payout_od_rate  = payout_val
#         payout_tp_rate  = payout_val
    
#     # CC Band
#     cc_band = row.get('CC Band', '')
#     from_cc, to_cc, _ = parse_cc_band(cc_band)
    
#     # is_cc_considered from Original Segment
#     is_cc_considered = get_is_cc_considered(row.get('Original Segment', ''))
    
#     # TW Type -> is_geared_vehicle & vehicle_type_id
#     tw_type = str(row.get('TW Type', '')).strip()
#     if 'scooter' in tw_type.lower():
#         is_geared = 0
#     else:
#         is_geared = 1  # Bikes
    
#     vt_id = get_vehicle_type_id(lob, tw_type)
    
#     out_row = {
#         'id': 0,
#         'company_id': comp_id,
#         'company_code': comp_code,
#         'segment_id': seg_id,
#         'segment': seg_name,
#         'subproduct_id': sub_prod_id,
#         'sub_product_name': sub_prod_name,
#         'lob_id': -1,
#         'lob_name': '',
#         'business_type_id': -1,
#         'business_type': 'Not Considered',
#         'is_highend_lob': False,
#         'rto_group_id': rto_group_id,
#         'rto_group_name': rto_group_name,
#         'payin_od_rate': payin_od_rate,
#         'payin_tp_rate': payin_tp_rate,
#         'payout_od_rate': payout_od_rate,
#         'payout_tp_rate': payout_tp_rate,
#         'extra_tp_rate': 0,
#         'eff_from_date': '2026-01-01',
#         'eff_to_date': '2026-01-16',
#         'fuel_type_id': -1,
#         'fuel_type': '',
#         'is_on_net': False,
#         'is_one_year_pay_on_newbusiness': -1,
#         'is_cpa_included': -1,
#         'is_geared_vehicle': is_geared,
#         'is_cc_considered': is_cc_considered,
#         'from_cc': from_cc,
#         'to_cc': to_cc,
#         'is_premium_considered': -1,
#         'from_premium': -1,
#         'to_premium': -1,
#         'is_mmv_considered': -1,
#         'make_id': -1,
#         'vehicle_make': '',
#         'model_id': -1,
#         'vehicle_model': '',
#         'variant_id': -1,
#         'vehicle_variant': '',
#         'is_seating_cap_consider': -1,
#         'from_seating_cap': -1,
#         'to_seating_cap': -1,
#         'is_no_of_wheel_consider': -1,
#         'from_no_of_wheel': -1,
#         'to_no_of_wheel': -1,
#         'vehicle_type_id': vt_id,
#         'ppi_in': 0,
#         'ppi_out': 0,
#         'is_irda_tp_included': -1,
#         'is_longterm_renewal_pay': -1,
#         'is_weightage_considered': -1,
#         'from_weightage_kg': 0,
#         'to_weightage_kg': 99999,
#         'is_nil_dep_considered': -1,
#         'is_organization_type': -1,
#         'from_age_month': 0,
#         'to_age_month': 700,
#         'is_with_ncb': -1,
#         'is_idv_cap_consider': -1,
#         'from_idv': 0,
#         'to_idv': 0,
#         'is_breakin_consider': -1,
#         'is_active': True,
#     }
#     rows.append(out_row)

# out_df = pd.DataFrame(rows)

# # Define column order as specified
# col_order = [
#     'id', 'company_id', 'company_code', 'segment_id', 'segment', 'subproduct_id', 'sub_product_name',
#     'lob_id', 'lob_name', 'business_type_id', 'business_type', 'is_highend_lob',
#     'rto_group_id', 'rto_group_name', 'payin_od_rate', 'payin_tp_rate', 'payout_od_rate', 'payout_tp_rate',
#     'extra_tp_rate', 'eff_from_date', 'eff_to_date', 'fuel_type_id', 'fuel_type',
#     'is_on_net', 'is_one_year_pay_on_newbusiness', 'is_cpa_included', 'is_geared_vehicle',
#     'is_cc_considered', 'from_cc', 'to_cc', 'is_premium_considered', 'from_premium', 'to_premium',
#     'is_mmv_considered', 'make_id', 'vehicle_make', 'model_id', 'vehicle_model', 'variant_id', 'vehicle_variant',
#     'is_seating_cap_consider', 'from_seating_cap', 'to_seating_cap',
#     'is_no_of_wheel_consider', 'from_no_of_wheel', 'to_no_of_wheel',
#     'vehicle_type_id', 'ppi_in', 'ppi_out',
#     'is_irda_tp_included', 'is_longterm_renewal_pay', 'is_weightage_considered',
#     'from_weightage_kg', 'to_weightage_kg', 'is_nil_dep_considered', 'is_organization_type',
#     'from_age_month', 'to_age_month', 'is_with_ncb', 'is_idv_cap_consider',
#     'from_idv', 'to_idv', 'is_breakin_consider', 'is_active'
# ]
# out_df = out_df[col_order]

# output_path = r"C:\Users\suyash.kamath\Downloads\TW_Processed_TW Comp_R_1_1771413426366.xlsx"
# out_df.to_excel(output_path, index=False)
# print(f"\n✓ Output saved: {output_path}")
# print(f"  Total records: {len(out_df)}")
# print("\nSample (first 3 rows):")
# print(out_df[['company_code', 'segment', 'sub_product_name', 'rto_group_name', 'payin_tp_rate', 'payout_tp_rate', 'from_cc', 'to_cc', 'vehicle_type_id']].head(3).to_string())

import pandas as pd
import json
import re
import sys
import os

# ─── Load JSON reference data ───────────────────────────────────────────────
def load_json(path):
    with open(path) as f:
        return json.load(f)

# ─── Ask for paths in CMD ────────────────────────────────────────────────────
print("\n" + "="*80)
print("PATH SETUP")
print("="*80)
BASE          = input("Enter JSON files folder path : ").strip().strip('"')
input_file    = input("Enter input Excel file path  : ").strip().strip('"')
output_folder = input("Enter output folder path     : ").strip().strip('"')

companies     = load_json(os.path.join(BASE, 'company_master.json'))
segments      = load_json(os.path.join(BASE, 'segment.json'))
subproducts   = load_json(os.path.join(BASE, 'subproduct.json'))
fuel_types    = load_json(os.path.join(BASE, 'fuel.json'))
vehicle_types = load_json(os.path.join(BASE, 'vehicle_type.json'))
rto_list      = load_json(os.path.join(BASE, 'rto_id_name.json'))

# Build lookup dicts
company_dict    = {c['company_id']: c for c in companies}
segment_dict    = {'Comprehensive': (1, 'Comprehensive'), 'SAOD': (2, 'SAOD'), 'TP Only': (3, 'TP Only')}
policy_map      = {'COMP': 'Comprehensive', 'TP': 'TP Only', 'SAOD': 'SAOD'}
subprod_dict    = {s['sub_product_name']: s['sub_product_id'] for s in subproducts}
lob_map         = {'TW': 'Two Wheeler', 'PC': 'Private Car', 'GCV': 'Goods Vehicle', 'PCV': 'Passenger Vehicle'}
vt_dict         = {(vt['sub_product_name'], vt['vehicle_type']): vt['id'] for vt in vehicle_types}
rto_dict        = {r['name']: r['id'] for r in rto_list}

# ─── Company selection ───────────────────────────────────────────────────────
print("\n" + "="*80)
print("AVAILABLE COMPANIES")
print("="*80)
for c in companies:
    print(f"  ID: {c['company_id']:3d} | Code: {c['company_code']:20s} | {c['company_name']}")
print("="*80)

company_id_input = input("\nEnter company_id from the list above: ").strip()
try:
    company_id = int(company_id_input)
    company = company_dict[company_id]
    print(f"\n✓ Selected: {company['company_name']} (ID: {company_id}, Code: {company['company_code']})")
except (ValueError, KeyError):
    print(f"ERROR: Invalid company_id '{company_id_input}'")
    sys.exit(1)

# ─── CC Band parser ──────────────────────────────────────────────────────────
def parse_cc_band(cc_band):
    if not cc_band or str(cc_band).strip() == '' or str(cc_band).strip().lower() == 'nan':
        return 0, 99999, -1
    
    s = str(cc_band).strip().upper().replace('CC', '').strip()
    
    m = re.match(r'^<\s*(\d+)$', s)
    if m:
        val = int(m.group(1))
        return 0, val - 1, 1
    
    m = re.match(r'^>\s*(\d+)$', s)
    if m:
        val = int(m.group(1))
        return val + 1, 99999, 1
    
    m = re.match(r'^>\s*(\d+)\s*[-–]\s*(\d+)$', s)
    if m:
        lo, hi = int(m.group(1)), int(m.group(2))
        return lo + 1, hi, 1
    
    m = re.match(r'^>=\s*(\d+)\s*[-–]\s*(\d+)$', s)
    if m:
        lo, hi = int(m.group(1)), int(m.group(2))
        return lo, hi, 1
    
    m = re.match(r'^(\d+)\s*[-–]\s*(\d+)$', s)
    if m:
        lo, hi = int(m.group(1)), int(m.group(2))
        return lo, hi, 1
    
    m = re.match(r'^(\d+)$', s)
    if m:
        val = int(m.group(1))
        return val, val, 1
    
    return 0, 99999, -1

# ─── Vehicle type ID resolver ────────────────────────────────────────────────
def get_vehicle_type_id(lob, tw_type):
    sub_prod = lob_map.get(lob, lob)
    if sub_prod == 'Two Wheeler':
        if tw_type and 'bike' in str(tw_type).lower():
            return 18  # TW Bike
        elif tw_type and 'scooter' in str(tw_type).lower():
            return 17  # TW Scooter
    return -1

# ─── is_cc_considered from Original Segment ──────────────────────────────────
def get_is_cc_considered(original_segment):
    if not original_segment or str(original_segment).strip().lower() == 'nan':
        return -1
    if re.search(r'\d', str(original_segment)):
        return 1
    return -1

# ─── Main processing ─────────────────────────────────────────────────────────
df = pd.read_excel(input_file)
print(f"\nProcessing {len(df)} records from input file...")

rows = []
for idx, row in df.iterrows():
    comp_id   = company_id
    comp_code = company['company_code']
    
    policy_type  = str(row.get('Policy Type', '')).strip()
    seg_name_raw = policy_map.get(policy_type, 'TP Only')
    seg_id, seg_name = segment_dict.get(seg_name_raw, (3, 'TP Only'))
    
    lob           = str(row.get('LOB', 'TW')).strip()
    sub_prod_name = lob_map.get(lob, lob)
    sub_prod_id   = subprod_dict.get(sub_prod_name, -1)
    
    # geo_new        = str(row.get('Geo New', '')).strip()
    # rto_group_id   = rto_dict.get(geo_new, -1)
   
    # rto_group_name = geo_new
    geo_new        = str(row.get('Geo Location', row.get('Geo New', ''))).strip()
    rto_group_id   = 0
    rto_group_name = geo_new
    
    payin_val  = row.get('Payin', 0)
    payout_val = row.get('Calculated Payout', 0)
    
    def to_float(v):
        if isinstance(v, str):
            v = v.strip().replace('%', '')
            try:
                return float(v) 
            except:
                return 0.0
        try:
            return float(v)
        except:
            return 0.0
    
    payin_val  = to_float(payin_val)
    payout_val = to_float(payout_val)
    
    if policy_type == 'TP':
        payin_od_rate, payin_tp_rate   = 0, payin_val
        payout_od_rate, payout_tp_rate = 0, payout_val
    elif policy_type == 'SAOD':
        payin_od_rate, payin_tp_rate   = payin_val, 0
        payout_od_rate, payout_tp_rate = payout_val, 0
    else:  # COMP
        payin_od_rate, payin_tp_rate   = payin_val, payin_val
        payout_od_rate, payout_tp_rate = payout_val, payout_val
    
    cc_band = row.get('CC Band', '')
    from_cc, to_cc, _ = parse_cc_band(cc_band)
    is_cc_considered  = get_is_cc_considered(row.get('Original Segment', ''))
    
    tw_type   = str(row.get('TW Type', '')).strip()
    is_geared = 0 if 'scooter' in tw_type.lower() else 1
    vt_id     = get_vehicle_type_id(lob, tw_type)
    
    rows.append({
        'id': 0,
        'company_id': comp_id,
        'company_code': comp_code,
        'segment_id': seg_id,
        'segment': seg_name,
        'subproduct_id': sub_prod_id,
        'sub_product_name': sub_prod_name,
        'lob_id': -1,
        'lob_name': '',
        'business_type_id': -1,
        'business_type': 'Not Considered',
        'is_highend_lob': False,
        'rto_group_id': rto_group_id,
        'rto_group_name': rto_group_name,
        'payin_od_rate': payin_od_rate,
        'payin_tp_rate': payin_tp_rate,
        'payout_od_rate': payout_od_rate,
        'payout_tp_rate': payout_tp_rate,
        'extra_tp_rate': 0,
        'eff_from_date': '2026-01-01',
        'eff_to_date': '2026-01-16',
        'fuel_type_id': -1,
        'fuel_type': '',
        'is_on_net': False,
        'is_one_year_pay_on_newbusiness': -1,
        'is_cpa_included': -1,
        'is_geared_vehicle': is_geared,
        'is_cc_considered': is_cc_considered,
        'from_cc': from_cc,
        'to_cc': to_cc,
        'is_premium_considered': -1,
        'from_premium': -1,
        'to_premium': -1,
        'is_mmv_considered': -1,
        'make_id': -1,
        'vehicle_make': '',
        'model_id': -1,
        'vehicle_model': '',
        'variant_id': -1,
        'vehicle_variant': '',
        'is_seating_cap_consider': -1,
        'from_seating_cap': -1,
        'to_seating_cap': -1,
        'is_no_of_wheel_consider': -1,
        'from_no_of_wheel': -1,
        'to_no_of_wheel': -1,
        'vehicle_type_id': vt_id,
        'ppi_in': 0,
        'ppi_out': 0,
        'is_irda_tp_included': -1,
        'is_longterm_renewal_pay': -1,
        'is_weightage_considered': -1,
        'from_weightage_kg': 0,
        'to_weightage_kg': 99999,
        'is_nil_dep_considered': -1,
        'is_organization_type': -1,
        'from_age_month': 0,
        'to_age_month': 700,
        'is_with_ncb': -1,
        'is_idv_cap_consider': -1,
        'from_idv': 0,
        'to_idv': 0,
        'is_breakin_consider': -1,
        'is_active': True,
    })

out_df = pd.DataFrame(rows)

col_order = [
    'id', 'company_id', 'company_code', 'segment_id', 'segment', 'subproduct_id', 'sub_product_name',
    'lob_id', 'lob_name', 'business_type_id', 'business_type', 'is_highend_lob',
    'rto_group_id', 'rto_group_name', 'payin_od_rate', 'payin_tp_rate', 'payout_od_rate', 'payout_tp_rate',
    'extra_tp_rate', 'eff_from_date', 'eff_to_date', 'fuel_type_id', 'fuel_type',
    'is_on_net', 'is_one_year_pay_on_newbusiness', 'is_cpa_included', 'is_geared_vehicle',
    'is_cc_considered', 'from_cc', 'to_cc', 'is_premium_considered', 'from_premium', 'to_premium',
    'is_mmv_considered', 'make_id', 'vehicle_make', 'model_id', 'vehicle_model', 'variant_id', 'vehicle_variant',
    'is_seating_cap_consider', 'from_seating_cap', 'to_seating_cap',
    'is_no_of_wheel_consider', 'from_no_of_wheel', 'to_no_of_wheel',
    'vehicle_type_id', 'ppi_in', 'ppi_out',
    'is_irda_tp_included', 'is_longterm_renewal_pay', 'is_weightage_considered',
    'from_weightage_kg', 'to_weightage_kg', 'is_nil_dep_considered', 'is_organization_type',
    'from_age_month', 'to_age_month', 'is_with_ncb', 'is_idv_cap_consider',
    'from_idv', 'to_idv', 'is_breakin_consider', 'is_active'
]
out_df = out_df[col_order]

output_path = os.path.join(output_folder, f'PayinConfig_{comp_code}_TW.xlsx')
out_df.to_excel(output_path, index=False)
print(f"\n✓ Output saved: {output_path}")
print(f"  Total records: {len(out_df)}")
print("\nSample (first 3 rows):")
print(out_df[['company_code', 'segment', 'sub_product_name', 'rto_group_name', 'payin_tp_rate', 'payout_tp_rate', 'from_cc', 'to_cc', 'vehicle_type_id']].head(3).to_string())
