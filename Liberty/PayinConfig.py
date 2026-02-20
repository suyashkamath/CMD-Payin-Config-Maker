# import pandas as pd
# import json
# import re
# import sys
# import os

# # ─── Load JSON reference data ───────────────────────────────────────────────
# def load_json(path):
#     with open(path) as f:
#         return json.load(f)

# # ─── Ask for paths in CMD ────────────────────────────────────────────────────
# print("\n" + "="*80)
# print("PATH SETUP")
# print("="*80)
# BASE          = input("Enter JSON files folder path : ").strip().strip('"')
# input_file    = input("Enter input Excel file path  : ").strip().strip('"')
# output_folder = input("Enter output folder path     : ").strip().strip('"')

# companies     = load_json(os.path.join(BASE, 'company_master.json'))
# segments      = load_json(os.path.join(BASE, 'segment.json'))
# subproducts   = load_json(os.path.join(BASE, 'subproduct.json'))
# fuel_types    = load_json(os.path.join(BASE, 'fuel.json'))
# vehicle_types = load_json(os.path.join(BASE, 'vehicle_type.json'))
# rto_list      = load_json(os.path.join(BASE, 'rto_id_name.json'))

# # Build lookup dicts
# company_dict    = {c['company_id']: c for c in companies}
# segment_dict    = {'Comprehensive': (1, 'Comprehensive'), 'SAOD': (2, 'SAOD'), 'TP Only': (3, 'TP Only')}
# policy_map      = {'COMP': 'Comprehensive', 'TP': 'TP Only', 'SAOD': 'SAOD'}
# subprod_dict    = {s['sub_product_name']: s['sub_product_id'] for s in subproducts}
# lob_map         = {'TW': 'Two Wheeler', 'PC': 'Private Car', 'GCV': 'Goods Vehicle', 'PCV': 'Passenger Vehicle'}
# vt_dict         = {(vt['sub_product_name'], vt['vehicle_type']): vt['id'] for vt in vehicle_types}
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
#     if not cc_band or str(cc_band).strip() == '' or str(cc_band).strip().lower() == 'nan':
#         return 0, 99999, -1
    
#     s = str(cc_band).strip().upper().replace('CC', '').strip()
    
#     m = re.match(r'^<\s*(\d+)$', s)
#     if m:
#         val = int(m.group(1))
#         return 0, val - 1, 1
    
#     m = re.match(r'^>\s*(\d+)$', s)
#     if m:
#         val = int(m.group(1))
#         return val + 1, 99999, 1
    
#     m = re.match(r'^>\s*(\d+)\s*[-–]\s*(\d+)$', s)
#     if m:
#         lo, hi = int(m.group(1)), int(m.group(2))
#         return lo + 1, hi, 1
    
#     m = re.match(r'^>=\s*(\d+)\s*[-–]\s*(\d+)$', s)
#     if m:
#         lo, hi = int(m.group(1)), int(m.group(2))
#         return lo, hi, 1
    
#     m = re.match(r'^(\d+)\s*[-–]\s*(\d+)$', s)
#     if m:
#         lo, hi = int(m.group(1)), int(m.group(2))
#         return lo, hi, 1
    
#     m = re.match(r'^(\d+)$', s)
#     if m:
#         val = int(m.group(1))
#         return val, val, 1
    
#     return 0, 99999, -1

# # ─── Vehicle type ID resolver ────────────────────────────────────────────────
# def get_vehicle_type_id(lob, tw_type):
#     sub_prod = lob_map.get(lob, lob)
#     if sub_prod == 'Two Wheeler':
#         if tw_type and 'bike' in str(tw_type).lower():
#             return  18  # TW Bike
#         elif tw_type and 'scooter' in str(tw_type).lower():
#             return 17  # TW Scooter
#     return -1

# # ─── is_cc_considered from Original Segment ──────────────────────────────────
# def get_is_cc_considered(original_segment):
#     if not original_segment or str(original_segment).strip().lower() == 'nan':
#         return -1
#     if re.search(r'\d', str(original_segment)):
#         return 1
#     return -1

# # ─── Main processing ─────────────────────────────────────────────────────────
# df = pd.read_excel(input_file)
# print(f"\nProcessing {len(df)} records from input file...")

# rows = []
# for idx, row in df.iterrows():
#     comp_id   = company_id
#     comp_code = company['company_code']
    
#     policy_type  = str(row.get('Policy Type', '')).strip()
#     seg_name_raw = policy_map.get(policy_type, 'TP Only')
#     seg_id, seg_name = segment_dict.get(seg_name_raw, (3, 'TP Only'))
    
#     lob           = str(row.get('LOB', 'TW')).strip()
#     sub_prod_name = lob_map.get(lob, lob)
#     sub_prod_id   = subprod_dict.get(sub_prod_name, -1)
    
#     # geo_new        = str(row.get('Geo New', '')).strip()
#     # rto_group_id   = rto_dict.get(geo_new, -1)
   
#     # rto_group_name = geo_new
#     geo_new        = str(row.get('Geo Location', row.get('Geo New', ''))).strip()
#     rto_group_id   = 0
#     rto_group_name = geo_new
    
#     payin_val  = row.get('Payin', 0)
#     payout_val = row.get('Calculated Payout', 0)
    
#     def to_float(v):
#         if isinstance(v, str):
#             v = v.strip().replace('%', '')
#             try:
#                 return float(v)
#             except:
#                 return 0.0
#         try:
#             return float(v)
#         except:
#             return 0.0
    
#     payin_val  = to_float(payin_val)
#     payout_val = to_float(payout_val)
    
#     if policy_type == 'TP':
#         payin_od_rate, payin_tp_rate   = 0, payin_val
#         payout_od_rate, payout_tp_rate = 0, payout_val
#     elif policy_type == 'SAOD':
#         payin_od_rate, payin_tp_rate   = payin_val, 0
#         payout_od_rate, payout_tp_rate = payout_val, 0
#     else:  # COMP
#         payin_od_rate, payin_tp_rate   = payin_val, payin_val
#         payout_od_rate, payout_tp_rate = payout_val, payout_val
    
#     cc_band = row.get('CC Band', '')
#     from_cc, to_cc, _ = parse_cc_band(cc_band)
#     is_cc_considered  = get_is_cc_considered(row.get('Original Segment', ''))
    
#     tw_type   = str(row.get('TW Type', '')).strip()
#     is_geared = 0 if 'scooter' in tw_type.lower() else 1
#     vt_id     = get_vehicle_type_id(lob, tw_type)
    
#     rows.append({
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
#         # 'is_on_net': False,
#         'is_on_net': True if policy_type == 'COMP' else False,
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
#     })

# out_df = pd.DataFrame(rows)

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

# # output_path = os.path.join(output_folder, f'PayinConfig_{comp_code}_TW.xlsx')
# # out_df.to_excel(output_path, index=False)
# output_path = os.path.join(output_folder, f'PayinConfig_{comp_code}_TW.xlsx')

# if os.path.exists(output_path):
#     existing_df = pd.read_excel(output_path)
#     out_df = pd.concat([existing_df, out_df], ignore_index=True)
#     print(f"\n✓ Appended to existing file!")

# out_df.to_excel(output_path, index=False)

# print(f"\n✓ Output saved: {output_path}")
# print(f"  Total records: {len(out_df)}")
# print("\nSample (first 3 rows):")
# print(out_df[['company_code', 'segment', 'sub_product_name', 'rto_group_name', 'payin_tp_rate', 'payout_tp_rate', 'from_cc', 'to_cc', 'vehicle_type_id']].head(3).to_string())


# import pandas as pd
# import json
# import re
# import sys
# import os

# # ─── Load JSON reference data ───────────────────────────────────────────────
# def load_json(path):
#     with open(path) as f:
#         return json.load(f)

# # ─── Ask for paths in CMD ────────────────────────────────────────────────────
# print("\n" + "="*80)
# print("PATH SETUP")
# print("="*80)
# BASE          = input("Enter JSON files folder path : ").strip().strip('"')
# output_folder = input("Enter output folder path     : ").strip().strip('"')

# companies     = load_json(os.path.join(BASE, 'company_master.json'))
# subproducts   = load_json(os.path.join(BASE, 'subproduct.json'))
# vehicle_types = load_json(os.path.join(BASE, 'vehicle_type.json'))
# rto_list      = load_json(os.path.join(BASE, 'rto_id_name.json'))

# # Build lookup dicts
# company_dict = {c['company_id']: c for c in companies}
# segment_dict = {'Comprehensive': (1, 'Comprehensive'), 'SAOD': (2, 'SAOD'), 'TP Only': (3, 'TP Only')}
# policy_map   = {'COMP': 'Comprehensive', 'TP': 'TP Only', 'SAOD': 'SAOD'}
# subprod_dict = {s['sub_product_name']: s['sub_product_id'] for s in subproducts}
# lob_map      = {'TW': 'Two Wheeler', 'PC': 'Private Car', 'GCV': 'Goods Vehicle', 'PCV': 'Passenger Vehicle', 'CV': 'Goods Vehicle'}
# vt_name_dict = {vt['vehicle_type']: vt['id'] for vt in vehicle_types}
# rto_dict     = {r['name']: r['id'] for r in rto_list}

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
#     company    = company_dict[company_id]
#     comp_code  = company['company_code']
#     print(f"\n✓ Selected: {company['company_name']} (ID: {company_id}, Code: {comp_code})")
# except (ValueError, KeyError):
#     print(f"ERROR: Invalid company_id '{company_id_input}'")
#     sys.exit(1)

# # ─── Output file path (single file for all runs) ────────────────────────────
# output_path = os.path.join(output_folder, f'{comp_code}-Payin-Config.xlsx')

# # ─── Helper: to_float ───────────────────────────────────────────────────────
# def to_float(v):
#     if isinstance(v, str):
#         v = v.strip().replace('%', '')
#         try:
#             return float(v)
#         except:
#             return 0.0
#     try:
#         return float(v)
#     except:
#         return 0.0

# # ─── CC Band parser (TW) ────────────────────────────────────────────────────
# def parse_cc_band(cc_band):
#     if not cc_band or str(cc_band).strip() == '' or str(cc_band).strip().lower() == 'nan':
#         return 0, 99999, -1
#     s = str(cc_band).strip().upper().replace('CC', '').strip()
#     m = re.match(r'^<\s*(\d+)$', s)
#     if m: return 0, int(m.group(1)) - 1, 1
#     m = re.match(r'^>\s*(\d+)$', s)
#     if m: return int(m.group(1)) + 1, 99999, 1
#     m = re.match(r'^>\s*(\d+)\s*[-–]\s*(\d+)$', s)
#     if m: return int(m.group(1)) + 1, int(m.group(2)), 1
#     m = re.match(r'^>=\s*(\d+)\s*[-–]\s*(\d+)$', s)
#     if m: return int(m.group(1)), int(m.group(2)), 1
#     m = re.match(r'^(\d+)\s*[-–]\s*(\d+)$', s)
#     if m: return int(m.group(1)), int(m.group(2)), 1
#     m = re.match(r'^(\d+)$', s)
#     if m: return int(m.group(1)), int(m.group(1)), 1
#     return 0, 99999, -1

# # ─── TW Vehicle type resolver ────────────────────────────────────────────────
# def get_tw_vehicle_type_id(tw_type):
#     if tw_type and 'bike' in str(tw_type).lower():
#         return 18  # TW Bike
#     elif tw_type and 'scooter' in str(tw_type).lower():
#         return 17  # TW Scooter
#     return -1

# # ─── is_cc_considered (TW) ──────────────────────────────────────────────────
# def get_is_cc_considered(original_segment):
#     if not original_segment or str(original_segment).strip().lower() == 'nan':
#         return -1
#     return 1 if re.search(r'\d', str(original_segment)) else -1

# # ─── Weight parser (CV) ─────────────────────────────────────────────────────
# def parse_weight(original_segment):
#     s = str(original_segment).strip()
#     m = re.search(r'[Uu]pto\s+([\d.]+)\s*[Tt]on?', s)
#     if m: return 1, 0, int(float(m.group(1)) * 1000)
#     m = re.search(r'>\s*([\d.]+)\s*~\s*([\d.]+)\s*[Tt]', s)
#     if m: return 1, int(float(m.group(1)) * 1000) + 100, int(float(m.group(2)) * 1000)
#     m = re.search(r'>\s*([\d.]+)\s*[Tt]', s)
#     if m: return 1, int(float(m.group(1)) * 1000) + 100, 99999
#     return -1, 0, 99999

# # ─── CV Vehicle info resolver ────────────────────────────────────────────────
# def get_cv_vehicle_info(original_segment):
#     s = str(original_segment).strip().upper()
#     if 'PCV 3W' in s:
#         return vt_name_dict.get('Auto rikshaw', -1), 'Passenger Vehicle', subprod_dict.get('Passenger Vehicle', -1), 1, 0, 3
#     if 'GCV 3W' in s:
#         return vt_name_dict.get('GCV 3W Delivery Van', -1), 'Goods Vehicle', subprod_dict.get('Goods Vehicle', -1), 1, 0, 3
#     return vt_name_dict.get('Truck', -1), 'Goods Vehicle', subprod_dict.get('Goods Vehicle', -1), -1, -1, -1

# # ─── Process a single file ───────────────────────────────────────────────────
# def process_file(input_file):
#     df = pd.read_excel(input_file)
#     print(f"\nProcessing {len(df)} records from: {os.path.basename(input_file)}")

#     lob_values = df['LOB'].dropna().unique().tolist() if 'LOB' in df.columns else []
#     detected   = 'CV' if any(str(l).strip().upper() == 'CV' for l in lob_values) else 'TW'
#     print(f"  Detected LOB type: {detected}")

#     rows = []
#     for idx, row in df.iterrows():

#         # Policy Type
#         policy_type = str(row.get('Policy Type', 'TP')).strip()
#         if policy_type in ('nan', ''):
#             policy_type = 'TP'
#         seg_name_raw     = policy_map.get(policy_type, 'TP Only')
#         seg_id, seg_name = segment_dict.get(seg_name_raw, (3, 'TP Only'))

#         lob              = str(row.get('LOB', 'TW')).strip().upper()
#         original_segment = str(row.get('Original Segment', '')).strip()

#         # RTO
#         geo_new        = str(row.get('Geo Location', row.get('Geo New', ''))).strip()
#         rto_group_id   = 0
#         rto_group_name = geo_new

#         # Payin / Payout
#         payin_val  = to_float(row.get('Payin', 0))
#         payout_val = to_float(row.get('Calculated Payout', 0))

#         if policy_type == 'TP':
#             payin_od_rate, payin_tp_rate   = 0, payin_val
#             payout_od_rate, payout_tp_rate = 0, payout_val
#         elif policy_type == 'SAOD':
#             payin_od_rate, payin_tp_rate   = payin_val, 0
#             payout_od_rate, payout_tp_rate = payout_val, 0
#         else:  # COMP
#             payin_od_rate, payin_tp_rate   = payin_val, payin_val
#             payout_od_rate, payout_tp_rate = payout_val, payout_val

#         # ── TW specific ──────────────────────────────────────────────────────
#         if lob == 'TW':
#             sub_prod_name = 'Two Wheeler'
#             sub_prod_id   = subprod_dict.get(sub_prod_name, -1)
#             tw_type       = str(row.get('TW Type', '')).strip()
#             is_geared     = 0 if 'scooter' in tw_type.lower() else 1
#             vt_id         = get_tw_vehicle_type_id(tw_type)
#             from_cc, to_cc, _ = parse_cc_band(row.get('CC Band', ''))
#             is_cc_considered  = get_is_cc_considered(original_segment)
#             is_weightage_considered, from_weightage_kg, to_weightage_kg = -1, 0, 99999
#             is_no_of_wheel, from_wheel, to_wheel = -1, -1, -1

#         # ── CV specific ──────────────────────────────────────────────────────
#         elif lob == 'CV':
#             vt_id, sub_prod_name, sub_prod_id, is_no_of_wheel, from_wheel, to_wheel = get_cv_vehicle_info(original_segment)
#             is_weightage_considered, from_weightage_kg, to_weightage_kg = parse_weight(original_segment)
#             is_geared        = -1
#             is_cc_considered = -1
#             from_cc, to_cc   = 0, 99999

#         # ── Fallback ─────────────────────────────────────────────────────────
#         else:
#             sub_prod_name = lob_map.get(lob, lob)
#             sub_prod_id   = subprod_dict.get(sub_prod_name, -1)
#             vt_id         = -1
#             is_geared     = -1
#             is_cc_considered, from_cc, to_cc = -1, 0, 99999
#             is_weightage_considered, from_weightage_kg, to_weightage_kg = -1, 0, 99999
#             is_no_of_wheel, from_wheel, to_wheel = -1, -1, -1

#         rows.append({
#             'id': 0,
#             'company_id': company_id,
#             'company_code': comp_code,
#             'segment_id': seg_id,
#             'segment': seg_name,
#             'subproduct_id': sub_prod_id,
#             'sub_product_name': sub_prod_name,
#             'lob_id': -1,
#             'lob_name': '',
#             'business_type_id': -1,
#             'business_type': 'Not Considered',
#             'is_highend_lob': False,
#             'rto_group_id': rto_group_id,
#             'rto_group_name': rto_group_name,
#             'payin_od_rate': payin_od_rate,
#             'payin_tp_rate': payin_tp_rate,
#             'payout_od_rate': payout_od_rate,
#             'payout_tp_rate': payout_tp_rate,
#             'extra_tp_rate': 0,
#             'eff_from_date': '2026-01-01',
#             'eff_to_date': '2026-01-16',
#             'fuel_type_id': -1,
#             'fuel_type': '',
#             'is_on_net': True if policy_type == 'COMP' else False,
#             'is_one_year_pay_on_newbusiness': -1,
#             'is_cpa_included': -1,
#             'is_geared_vehicle': is_geared,
#             'is_cc_considered': is_cc_considered,
#             'from_cc': from_cc,
#             'to_cc': to_cc,
#             'is_premium_considered': -1,
#             'from_premium': -1,
#             'to_premium': -1,
#             'is_mmv_considered': -1,
#             'make_id': -1,
#             'vehicle_make': '',
#             'model_id': -1,
#             'vehicle_model': '',
#             'variant_id': -1,
#             'vehicle_variant': '',
#             'is_seating_cap_consider': -1,
#             'from_seating_cap': -1,
#             'to_seating_cap': -1,
#             'is_no_of_wheel_consider': is_no_of_wheel,
#             'from_no_of_wheel': from_wheel,
#             'to_no_of_wheel': to_wheel,
#             'vehicle_type_id': vt_id,
#             'ppi_in': 0,
#             'ppi_out': 0,
#             'is_irda_tp_included': -1,
#             'is_longterm_renewal_pay': -1,
#             'is_weightage_considered': is_weightage_considered,
#             'from_weightage_kg': from_weightage_kg,
#             'to_weightage_kg': to_weightage_kg,
#             'is_nil_dep_considered': -1,
#             'is_organization_type': -1,
#             'from_age_month': 0,
#             'to_age_month': 700,
#             'is_with_ncb': -1,
#             'is_idv_cap_consider': -1,
#             'from_idv': 0,
#             'to_idv': 0,
#             'is_breakin_consider': -1,
#             'is_active': True,
#         })

#     return pd.DataFrame(rows)

# # ─── Column order ────────────────────────────────────────────────────────────

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

# # ─── Main loop ───────────────────────────────────────────────────────────────
# input_file = input("\nEnter input Excel file path  : ").strip().strip('"')

# while True:
#     try:
#         out_df = process_file(input_file)
#         out_df = out_df[col_order]

#         if os.path.exists(output_path):
#             existing_df = pd.read_excel(output_path)
#             out_df = pd.concat([existing_df, out_df], ignore_index=True)
#             print(f"\n✓ Appended to existing file!")

#         out_df.to_excel(output_path, index=False)
#         print(f"✓ Output saved : {output_path}")
#         print(f"  Total records: {len(out_df)}")

#     except Exception as e:
#         print(f"\nERROR processing file: {e}")

#     print("\n" + "="*80)
#     print("  1. Yes - Add more files")
#     print("  2. No  - Exit")
#     choice = input("Do you want to add more files? Enter choice: ").strip()

#     if choice == '2':
#         print("\n✓ Done! Exiting.")
#         break
#     else:
#         input_file = input("\nEnter next input Excel file path : ").strip().strip('"')


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
output_folder = input("Enter output folder path     : ").strip().strip('"')

companies     = load_json(os.path.join(BASE, 'company_master.json'))
subproducts   = load_json(os.path.join(BASE, 'subproduct.json'))
vehicle_types = load_json(os.path.join(BASE, 'vehicle_type.json'))
rto_list      = load_json(os.path.join(BASE, 'rto_id_name.json'))
fuel_list     = load_json(os.path.join(BASE, 'fuel.json'))

# Build lookup dicts
company_dict  = {c['company_id']: c for c in companies}
segment_dict  = {'Comprehensive': (1, 'Comprehensive'), 'SAOD': (2, 'SAOD'), 'TP Only': (3, 'TP Only')}
policy_map    = {'COMP': 'Comprehensive', 'TP': 'TP Only', 'SAOD': 'SAOD'}
subprod_dict  = {s['sub_product_name']: s['sub_product_id'] for s in subproducts}
lob_map       = {'TW': 'Two Wheeler', 'PC': 'Private Car', 'PVT CAR': 'Private Car', 'GCV': 'Goods Vehicle', 'PCV': 'Passenger Vehicle', 'CV': 'Goods Vehicle'}
vt_name_dict  = {vt['vehicle_type']: vt['id'] for vt in vehicle_types}
rto_dict      = {r['name']: r['id'] for r in rto_list}
fuel_dict     = {f['fuel_type_name'].upper(): f['fuel_type_id'] for f in fuel_list}
# Others fuel types = all except PETROL
OTHERS_FUELS  = [
    (fuel_dict['DIESEL'],  'DIESEL'),
    (fuel_dict['ELECTRIC'], 'ELECTRIC'),
    (fuel_dict['CNG-LPG'], 'CNG-LPG'),
]

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
    company    = company_dict[company_id]
    comp_code  = company['company_code']
    print(f"\n✓ Selected: {company['company_name']} (ID: {company_id}, Code: {comp_code})")
except (ValueError, KeyError):
    print(f"ERROR: Invalid company_id '{company_id_input}'")
    sys.exit(1)

output_path = os.path.join(output_folder, f'{comp_code}-Payin-Config.xlsx')

# ─── Helpers ────────────────────────────────────────────────────────────────
def to_float(v):
    if isinstance(v, str):
        v = v.strip().replace('%', '')
        try: return float(v)
        except: return 0.0
    try: return float(v)
    except: return 0.0

def parse_cc_band(cc_band):
    if not cc_band or str(cc_band).strip() == '' or str(cc_band).strip().lower() == 'nan':
        return 0, 99999, -1
    s = str(cc_band).strip().upper().replace('CC', '').strip()
    m = re.match(r'^<\s*(\d+)$', s)
    if m: return 0, int(m.group(1)) - 1, 1
    m = re.match(r'^>\s*(\d+)$', s)
    if m: return int(m.group(1)) + 1, 99999, 1
    m = re.match(r'^>\s*(\d+)\s*[-–]\s*(\d+)$', s)
    if m: return int(m.group(1)) + 1, int(m.group(2)), 1
    m = re.match(r'^>=\s*(\d+)\s*[-–]\s*(\d+)$', s)
    if m: return int(m.group(1)), int(m.group(2)), 1
    m = re.match(r'^(\d+)\s*[-–]\s*(\d+)$', s)
    if m: return int(m.group(1)), int(m.group(2)), 1
    m = re.match(r'^(\d+)$', s)
    if m: return int(m.group(1)), int(m.group(1)), 1
    return 0, 99999, -1

def get_tw_vehicle_type_id(tw_type):
    if tw_type and 'bike' in str(tw_type).lower(): return 18
    elif tw_type and 'scooter' in str(tw_type).lower(): return 17
    return -1

def get_is_cc_considered(original_segment):
    if not original_segment or str(original_segment).strip().lower() == 'nan': return -1
    return 1 if re.search(r'\d', str(original_segment)) else -1

def parse_weight(original_segment):
    s = str(original_segment).strip()
    m = re.search(r'[Uu]pto\s+([\d.]+)\s*[Tt]on?', s)
    if m: return 1, 0, int(float(m.group(1)) * 1000)
    m = re.search(r'>\s*([\d.]+)\s*~\s*([\d.]+)\s*[Tt]', s)
    if m: return 1, int(float(m.group(1)) * 1000) + 100, int(float(m.group(2)) * 1000)
    m = re.search(r'>\s*([\d.]+)\s*[Tt]', s)
    if m: return 1, int(float(m.group(1)) * 1000) + 100, 99999
    return -1, 0, 99999

def get_cv_vehicle_info(original_segment):
    s = str(original_segment).strip().upper()
    if 'PCV 3W' in s:
        return vt_name_dict.get('Auto rikshaw', -1), 'Passenger Vehicle', subprod_dict.get('Passenger Vehicle', -1), 1, 0, 3
    if 'GCV 3W' in s:
        return vt_name_dict.get('GCV 3W Delivery Van', -1), 'Goods Vehicle', subprod_dict.get('Goods Vehicle', -1), 1, 0, 3
    return vt_name_dict.get('Truck', -1), 'Goods Vehicle', subprod_dict.get('Goods Vehicle', -1), -1, -1, -1

# ─── PC Original Segment parser ─────────────────────────────────────────────
# Returns list of (fuel_type_id, fuel_type_name, is_with_ncb) tuples
# One row may expand into multiple output rows
def parse_pc_segment(original_segment):
    """
    Examples:
      'Comp - Petrol - NCB / NON NCB'     -> [(2, PETROL, 1), (2, PETROL, 0)]
      'Comp - Diesel / Others - NCB'      -> [(1, DIESEL, 1), (3, ELECTRIC, 1), (4, CNG-LPG, 1)]
      'Comp - Diesel / Others - Non NCB'  -> [(1, DIESEL, 0), (3, ELECTRIC, 0), (4, CNG-LPG, 0)]
      'SOD - NCB'                         -> [(-1, '', 1)]
      'SOD - NON NCB'                     -> [(-1, '', 0)]
    """
    s = str(original_segment).strip().upper()

    # Determine NCB flags
    has_ncb     = 'NCB' in s and 'NON NCB' not in s.replace('NON NCB', '')
    has_non_ncb = 'NON NCB' in s
    has_both    = 'NCB / NON NCB' in s or 'NON NCB' in s and 'NCB' in s

    if 'NCB / NON NCB' in s:
        ncb_flags = [1, 0]
    elif 'NON NCB' in s:
        ncb_flags = [0]
    elif 'NCB' in s:
        ncb_flags = [1]
    else:
        ncb_flags = [-1]

    # Determine fuel types
    if 'PETROL' in s:
        fuels = [(fuel_dict['PETROL'], 'PETROL')]
    elif 'DIESEL' in s and 'OTHERS' in s:
        # Others = Diesel + Electric + CNG-LPG
        fuels = OTHERS_FUELS
    elif 'DIESEL' in s:
        fuels = [(fuel_dict['DIESEL'], 'DIESEL')]
    else:
        # SOD or no fuel mentioned
        fuels = [(-1, '')]

    # Expand: one entry per (fuel, ncb) combination
    result = []
    for fuel_id, fuel_name in fuels:
        for ncb in ncb_flags:
            result.append((fuel_id, fuel_name, ncb))
    return result

# ─── Build a single output row dict ─────────────────────────────────────────
def build_row(company_id, comp_code, seg_id, seg_name, sub_prod_id, sub_prod_name,
              rto_group_id, rto_group_name,
              payin_od_rate, payin_tp_rate, payout_od_rate, payout_tp_rate,
              policy_type, fuel_type_id, fuel_type_name, is_with_ncb,
              is_geared, is_cc_considered, from_cc, to_cc,
              is_weightage_considered, from_weightage_kg, to_weightage_kg,
              is_no_of_wheel, from_wheel, to_wheel, vt_id):
    return {
        'id': 0,
        'company_id': company_id,
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
        'fuel_type_id': fuel_type_id,
        'fuel_type': fuel_type_name,
        'is_on_net': True if policy_type == 'COMP' else False,
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
        'is_no_of_wheel_consider': is_no_of_wheel,
        'from_no_of_wheel': from_wheel,
        'to_no_of_wheel': to_wheel,
        'vehicle_type_id': vt_id,
        'ppi_in': 0,
        'ppi_out': 0,
        'is_irda_tp_included': -1,
        'is_longterm_renewal_pay': -1,
        'is_weightage_considered': is_weightage_considered,
        'from_weightage_kg': from_weightage_kg,
        'to_weightage_kg': to_weightage_kg,
        'is_nil_dep_considered': -1,
        'is_organization_type': -1,
        'from_age_month': 0,
        'to_age_month': 700,
        'is_with_ncb': is_with_ncb,
        'is_idv_cap_consider': -1,
        'from_idv': 0,
        'to_idv': 0,
        'is_breakin_consider': -1,
        'is_active': True,
    }

# ─── Process a single file ───────────────────────────────────────────────────
def process_file(input_file):
    df = pd.read_excel(input_file)
    print(f"\nProcessing {len(df)} records from: {os.path.basename(input_file)}")

    lob_values = df['LOB'].dropna().unique().tolist() if 'LOB' in df.columns else []
    lob_upper  = [str(l).strip().upper() for l in lob_values]

    is_tw  = any(l == 'TW' for l in lob_upper)
    is_cv  = any(l == 'CV' for l in lob_upper)
    is_pc  = any(l in ('PC', 'PVT CAR') for l in lob_upper)

    detected = 'TW' if is_tw else ('CV' if is_cv else 'PC')
    print(f"  Detected LOB type: {detected}")

    rows = []
    for idx, row in df.iterrows():

        # Policy Type
        policy_type = str(row.get('Policy Type', 'TP')).strip()
        if policy_type in ('nan', ''): policy_type = 'TP'
        seg_name_raw     = policy_map.get(policy_type, 'TP Only')
        seg_id, seg_name = segment_dict.get(seg_name_raw, (3, 'TP Only'))

        lob              = str(row.get('LOB', 'TW')).strip().upper()
        original_segment = str(row.get('Original Segment', '')).strip()
        geo_new          = str(row.get('Geo Location', row.get('Geo New', ''))).strip()
        rto_group_id     = 0
        rto_group_name   = geo_new

        # Payin column — PC uses 'Payin (OD Premium)', others use 'Payin'
        payin_col  = 'Payin (OD Premium)' if 'Payin (OD Premium)' in df.columns else 'Payin'
        payin_val  = to_float(row.get(payin_col, 0))
        payout_val = to_float(row.get('Calculated Payout', 0))

        if policy_type == 'TP':
            payin_od_rate, payin_tp_rate   = 0, payin_val
            payout_od_rate, payout_tp_rate = 0, payout_val
        elif policy_type == 'SAOD':
            payin_od_rate, payin_tp_rate   = payin_val, 0
            payout_od_rate, payout_tp_rate = payout_val, 0
        else:  # COMP
            payin_od_rate, payin_tp_rate   = payin_val, payin_val
            payout_od_rate, payout_tp_rate = payout_val, payout_val

        # ── TW ───────────────────────────────────────────────────────────────
        if lob == 'TW':
            sub_prod_name = 'Two Wheeler'
            sub_prod_id   = subprod_dict.get(sub_prod_name, -1)
            tw_type       = str(row.get('TW Type', '')).strip()
            is_geared     = 0 if 'scooter' in tw_type.lower() else 1
            vt_id         = get_tw_vehicle_type_id(tw_type)
            from_cc, to_cc, _ = parse_cc_band(row.get('CC Band', ''))
            is_cc_considered  = get_is_cc_considered(original_segment)
            rows.append(build_row(
                company_id, comp_code, seg_id, seg_name, sub_prod_id, sub_prod_name,
                rto_group_id, rto_group_name,
                payin_od_rate, payin_tp_rate, payout_od_rate, payout_tp_rate,
                policy_type, -1, '', -1,
                is_geared, is_cc_considered, from_cc, to_cc,
                -1, 0, 99999, -1, -1, -1, vt_id
            ))

        # ── CV ───────────────────────────────────────────────────────────────
        elif lob == 'CV':
            vt_id, sub_prod_name, sub_prod_id, is_no_of_wheel, from_wheel, to_wheel = get_cv_vehicle_info(original_segment)
            is_wt, from_wt, to_wt = parse_weight(original_segment)
            rows.append(build_row(
                company_id, comp_code, seg_id, seg_name, sub_prod_id, sub_prod_name,
                rto_group_id, rto_group_name,
                payin_od_rate, payin_tp_rate, payout_od_rate, payout_tp_rate,
                policy_type, -1, '', -1,
                -1, -1, 0, 99999,
                is_wt, from_wt, to_wt, is_no_of_wheel, from_wheel, to_wheel, vt_id
            ))

        # ── PC ───────────────────────────────────────────────────────────────
        elif lob in ('PC', 'PVT CAR'):
            sub_prod_name = 'Private Car'
            sub_prod_id   = subprod_dict.get(sub_prod_name, -1)
            vt_id         = vt_name_dict.get('Private Car', -1)

            # Expand row by fuel + NCB combinations
            expansions = parse_pc_segment(original_segment)
            for fuel_id, fuel_name, is_with_ncb in expansions:
                rows.append(build_row(
                    company_id, comp_code, seg_id, seg_name, sub_prod_id, sub_prod_name,
                    rto_group_id, rto_group_name,
                    payin_od_rate, payin_tp_rate, payout_od_rate, payout_tp_rate,
                    policy_type, fuel_id, fuel_name, is_with_ncb,
                    -1, -1, 0, 99999,
                    -1, 0, 99999, -1, -1, -1, vt_id
                ))

        # ── Fallback ─────────────────────────────────────────────────────────
        else:
            sub_prod_name = lob_map.get(lob, lob)
            sub_prod_id   = subprod_dict.get(sub_prod_name, -1)
            rows.append(build_row(
                company_id, comp_code, seg_id, seg_name, sub_prod_id, sub_prod_name,
                rto_group_id, rto_group_name,
                payin_od_rate, payin_tp_rate, payout_od_rate, payout_tp_rate,
                policy_type, -1, '', -1,
                -1, -1, 0, 99999,
                -1, 0, 99999, -1, -1, -1, -1
            ))

    print(f"  Expanded to {len(rows)} output rows")
    return pd.DataFrame(rows)

# ─── Column order ────────────────────────────────────────────────────────────
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

# ─── Main loop ───────────────────────────────────────────────────────────────
input_file = input("\nEnter input Excel file path  : ").strip().strip('"')

while True:
    try:
        out_df = process_file(input_file)
        out_df = out_df[col_order]

        if os.path.exists(output_path):
            existing_df = pd.read_excel(output_path)
            out_df = pd.concat([existing_df, out_df], ignore_index=True)
            print(f"\n✓ Appended to existing file!")

        out_df.to_excel(output_path, index=False)
        print(f"✓ Output saved : {output_path}")
        print(f"  Total records: {len(out_df)}")

    except Exception as e:
        print(f"\nERROR processing file: {e}")

    print("\n" + "="*80)
    print("  1. Yes - Add more files")
    print("  2. No  - Exit")
    choice = input("Do you want to add more files? Enter choice: ").strip()

    if choice == '2':
        print("\n✓ Done! Exiting.")
        break
    else:
        input_file = input("\nEnter next input Excel file path : ").strip().strip('"')
