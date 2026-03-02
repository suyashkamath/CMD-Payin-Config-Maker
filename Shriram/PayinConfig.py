# # # import pandas as pd
# # # import json
# # # import re
# # # import sys
# # # import os
# # # import urllib.request
# # # import urllib.error

# # # # ─── Load JSON reference data ───────────────────────────────────────────────
# # # def load_json(path):
# # #     with open(path) as f:
# # #         return json.load(f)

# # # def load_dotenv_if_exists(dotenv_path):
# # #     if not os.path.exists(dotenv_path):
# # #         return
# # #     with open(dotenv_path, encoding="utf-8") as f:
# # #         for line in f:
# # #             s = line.strip()
# # #             if not s or s.startswith("#") or "=" not in s:
# # #                 continue
# # #             k, v = s.split("=", 1)
# # #             k = k.strip()
# # #             v = v.strip().strip('"').strip("'")
# # #             if k and k not in os.environ:
# # #                 os.environ[k] = v

# # # def normalize_lookup_text(v):
# # #     s = str(v).upper()
# # #     s = re.sub(r'[^A-Z0-9]+', ' ', s)
# # #     return re.sub(r'\s+', ' ', s).strip()

# # # # ─── Ask for paths in CMD ────────────────────────────────────────────────────
# # # print("\n" + "="*80)
# # # print("PATH SETUP")
# # # print("="*80)
# # # BASE          = input("Enter JS files folder path : ").strip().strip('"')
# # # output_folder = input("Enter output folder path     : ").strip().strip('"')
# # # load_dotenv_if_exists(os.path.join(BASE, ".env"))
# # # load_dotenv_if_exists(".env")

# # # companies     = load_json(os.path.join(BASE, 'company_master.json'))
# # # subproducts   = load_json(os.path.join(BASE, 'subproduct.json'))
# # # vehicle_types = load_json(os.path.join(BASE, 'vehicle_type.json'))
# # # rto_list      = load_json(os.path.join(BASE, 'rto_id_name.json'))
# # # fuel_list     = load_json(os.path.join(BASE, 'fuel.json'))
# # # cv_mmv_path_1 = os.path.join(BASE, 'cv_mmv_master.json')
# # # cv_mmv_path_2 = 'cv_mmv_master.json'
# # # if os.path.exists(cv_mmv_path_1):
# # #     cv_mmv_list = load_json(cv_mmv_path_1)
# # # elif os.path.exists(cv_mmv_path_2):
# # #     cv_mmv_list = load_json(cv_mmv_path_2)
# # # else:
# # #     cv_mmv_list = []

# # # # Build lookup dicts
# # # company_dict = {c['company_id']: c for c in companies}
# # # segment_dict = {'Comprehensive': (1, 'Comprehensive'), 'SAOD': (2, 'SAOD'), 'TP Only': (3, 'TP Only')}
# # # policy_map   = {'COMP': 'Comprehensive', 'TP': 'TP Only', 'SAOD': 'SAOD'}
# # # subprod_dict = {s['sub_product_name']: s['sub_product_id'] for s in subproducts}
# # # lob_map      = {'TW': 'Two Wheeler', 'PC': 'Private Car', 'PVT CAR': 'Private Car', 'GCV': 'Goods Vehicle', 'PCV': 'Passenger Vehicle', 'CV': 'Goods Vehicle'}
# # # vt_name_dict = {vt['vehicle_type']: vt['id'] for vt in vehicle_types}
# # # rto_dict     = {r['name']: r['id'] for r in rto_list}
# # # fuel_dict    = {f['fuel_type_name'].upper(): f['fuel_type_id'] for f in fuel_list}
# # # OTHERS_FUELS = [
# # #     (fuel_dict['DIESEL'],   'DIESEL'),
# # #     (fuel_dict['ELECTRIC'], 'ELECTRIC'),
# # #     (fuel_dict['CNG-LPG'],  'CNG-LPG'),
# # # ]

# # # mmv_make_index = {}
# # # mmv_model_index = {}
# # # mmv_company_make_names = {}
# # # mmv_company_model_names = {}
# # # mmv_models_by_make = {}
# # # for m in cv_mmv_list:
# # #     code = normalize_lookup_text(m.get("company_code", ""))
# # #     make_name = str(m.get("vehicle_make", "")).strip()
# # #     model_name = str(m.get("vehicle_model", "")).strip()
# # #     make_norm = normalize_lookup_text(make_name)
# # #     model_norm = normalize_lookup_text(model_name)
# # #     make_id = int(m.get("MakeID", -1)) if str(m.get("MakeID", "")).strip() not in ("", "nan", "None") else -1
# # #     model_id = int(m.get("ModelID", -1)) if str(m.get("ModelID", "")).strip() not in ("", "nan", "None") else -1
# # #     if code and make_norm and make_id != -1:
# # #         mmv_make_index[(code, make_norm)] = (make_id, make_name)
# # #         mmv_company_make_names.setdefault(code, set()).add((make_norm, make_name, make_id))
# # #     if code and make_norm and model_norm and model_id != -1:
# # #         mmv_model_index[(code, make_norm, model_norm)] = (model_id, model_name)
# # #         mmv_company_model_names.setdefault(code, {}).setdefault(model_norm, []).append((model_id, model_name, make_id, make_name))
# # #         mmv_models_by_make.setdefault((code, make_norm), set()).add((model_norm, model_name, model_id))

# # # for k in list(mmv_company_make_names.keys()):
# # #     mmv_company_make_names[k] = sorted(
# # #         list(mmv_company_make_names[k]),
# # #         key=lambda x: len(x[0]),
# # #         reverse=True
# # #     )

# # # def keep_included_text(remark_text):
# # #     s = str(remark_text).strip()
# # #     if not s:
# # #         return ""
# # #     parts = re.split(r'[;,]', s)
# # #     kept = []
# # #     for p in parts:
# # #         up = p.upper()
# # #         # Hard negative chunks should be fully ignored.
# # #         if any(tok in up for tok in ("DECLIN", "REJECT")):
# # #             continue
# # #         # Keep included side for these connectors.
# # #         for token in (" BUT ", " EXCEPT ", " EXCLUDE ", " OTHER THAN ", " NOT CONSIDER"):
# # #             idx = up.find(token)
# # #             if idx != -1:
# # #                 p = p[:idx]
# # #                 break
# # #         p = p.strip()
# # #         if p:
# # #             kept.append(p)
# # #     return " ".join([x for x in kept if x])

# # # # ─── Company selection ───────────────────────────────────────────────────────
# # # print("\n" + "="*80)
# # # print("AVAILABLE COMPANIES")
# # # print("="*80)
# # # for c in companies:
# # #     print(f"  ID: {c['company_id']:3d} | Code: {c['company_code']:20s} | {c['company_name']}")
# # # print("="*80)

# # # company_id_input = input("\nEnter company_id from the list above: ").strip()
# # # try:
# # #     company_id = int(company_id_input)
# # #     company    = company_dict[company_id]
# # #     comp_code  = company['company_code']
# # #     print(f"\n✓ Selected: {company['company_name']} (ID: {company_id}, Code: {comp_code})")
# # # except (ValueError, KeyError):
# # #     print(f"ERROR: Invalid company_id '{company_id_input}'")
# # #     sys.exit(1)

# # # output_path = os.path.join(output_folder, f'{comp_code}-Payin-Config.xlsx')

# # # # ─── Helpers ────────────────────────────────────────────────────────────────
# # # def to_float(v):
# # #     if isinstance(v, str):
# # #         v = v.strip().replace('%', '')
# # #         try: return float(v)
# # #         except: return 0.0
# # #     try: return float(v)
# # #     except: return 0.0

# # # def parse_cc_band(cc_band):
# # #     if not cc_band or str(cc_band).strip() == '' or str(cc_band).strip().lower() == 'nan':
# # #         return 0, 99999, -1
# # #     s = str(cc_band).strip().upper().replace('CC', '').strip()
# # #     m = re.match(r'^<\s*(\d+)$', s)
# # #     if m: return 0, int(m.group(1)) - 1, 1
# # #     m = re.match(r'^>\s*(\d+)$', s)
# # #     if m: return int(m.group(1)) + 1, 99999, 1
# # #     m = re.match(r'^>\s*(\d+)\s*[-–]\s*(\d+)$', s)
# # #     if m: return int(m.group(1)) + 1, int(m.group(2)), 1
# # #     m = re.match(r'^>=\s*(\d+)\s*[-–]\s*(\d+)$', s)
# # #     if m: return int(m.group(1)), int(m.group(2)), 1
# # #     m = re.match(r'^(\d+)\s*[-–]\s*(\d+)$', s)
# # #     if m: return int(m.group(1)), int(m.group(2)), 1
# # #     m = re.match(r'^(\d+)$', s)
# # #     if m: return int(m.group(1)), int(m.group(1)), 1
# # #     return 0, 99999, -1

# # # def get_tw_vehicle_type_id(tw_type):
# # #     if tw_type and 'bike' in str(tw_type).lower(): return 18
# # #     elif tw_type and 'scooter' in str(tw_type).lower(): return 17
# # #     return -1

# # # def get_is_cc_considered(original_segment):
# # #     if not original_segment or str(original_segment).strip().lower() == 'nan': return -1
# # #     return 1 if re.search(r'\d', str(original_segment)) else -1

# # # def parse_weight(original_segment):
# # #     """Used for CV TP — parses weight from Original Segment text"""
# # #     s = str(original_segment).strip()
# # #     m = re.search(r'[Uu]pto\s+([\d.]+)\s*[Tt]on?', s)
# # #     if m: return 1, 0, int(float(m.group(1)) * 1000)
# # #     m = re.search(r'>\s*([\d.]+)\s*~\s*([\d.]+)\s*[Tt]', s)
# # #     if m: return 1, int(float(m.group(1)) * 1000) + 100, int(float(m.group(2)) * 1000)
# # #     m = re.search(r'>\s*([\d.]+)\s*[Tt]', s)
# # #     if m: return 1, int(float(m.group(1)) * 1000) + 100, 99999
# # #     return -1, 0, 99999

# # # def parse_tonnage_col(tonnage):
# # #     """Used for CV COMP — parses weight from Tonnage column"""
# # #     s = str(tonnage).strip().upper()
# # #     if s in ('NAN', ''):
# # #         return -1, 0, 99999
# # #     # PCV 3W or MISD Tractors — no weight range
# # #     if 'PCV' in s or 'MISD' in s or 'TRACTOR' in s:
# # #         return -1, 0, 99999
# # #     # Range: 0 - 2.6T, 4 - 7.5T, 12 - 20T, 20 - 43T
# # #     m = re.match(r'^([\d.]+)\s*[-–]\s*([\d.]+)\s*T', s)
# # #     if m:
# # #         lo = int(float(m.group(1)) * 1000)
# # #         hi = int(float(m.group(2)) * 1000)
# # #         return 1, lo, hi
# # #     # Single value like 7.5T+
# # #     m = re.match(r'^([\d.]+)\s*T', s)
# # #     if m:
# # #         return 1, int(float(m.group(1)) * 1000), 99999
# # #     return -1, 0, 99999

# # # def parse_age_band(age_band):
# # #     """
# # #     Converts age band string to (from_age_month, to_age_month)
# # #     New            -> 0 to 8
# # #     >1 - 5 Years   -> 13 to 60
# # #     >5 - 10+ Years -> 61 to 700  (+ means max = 700)
# # #     """
# # #     s = str(age_band).strip()
# # #     su = s.upper()

# # #     if su == 'NEW':
# # #         return 0, 8

# # #     # Pattern: >X - Y+ Years  or  >X - Y Years
# # #     m = re.match(r'>\s*(\d+)\s*[-–]\s*(\d+)(\+?)\s*[Yy]ear', s)
# # #     if m:
# # #         lo   = int(m.group(1)) * 12 + 1
# # #         hi_n = int(m.group(2))
# # #         has_plus = m.group(3) == '+'
# # #         hi = 700 if has_plus else hi_n * 12
# # #         return lo, hi

# # #     # Pattern: >X+ Years
# # #     m = re.match(r'>\s*(\d+)\+\s*[Yy]ear', s)
# # #     if m:
# # #         lo = int(m.group(1)) * 12 + 1
# # #         return lo, 700

# # #     # Pattern: >X Years (no upper)
# # #     m = re.match(r'>\s*(\d+)\s*[Yy]ear', s)
# # #     if m:
# # #         lo = int(m.group(1)) * 12 + 1
# # #         return lo, 700

# # #     return 0, 700

# # # def normalize_policy_type(policy_type):
# # #     s = str(policy_type).strip().upper()
# # #     if s in ("COMP", "COMPREHENSIVE"):
# # #         return "COMP"
# # #     if s in ("TP", "TP ONLY", "THIRD PARTY", "THIRD PARTY ONLY"):
# # #         return "TP"
# # #     if s in ("SAOD", "OD", "OWN DAMAGE"):
# # #         return "SAOD"
# # #     return "TP"

# # # def parse_shriram_age_band(age_text):
# # #     s = str(age_text).strip()
# # #     if s == "" or s.lower() == "nan":
# # #         return 0, 700
# # #     m = re.search(r'UPTO\s+(\d+)', s.upper())
# # #     if m:
# # #         return 0, int(m.group(1)) * 12
# # #     return parse_age_band(s)

# # # def get_shriram_vehicle_info(segment_text):
# # #     s = str(segment_text).strip().upper()
# # #     # User rule: "All GVW & PCV 3W, GCV 3W" should be treated as Goods Vehicle.
# # #     if "ALL GVW & PCV 3W, GCV 3W" in s:
# # #         return "CV", "Goods Vehicle", subprod_dict.get("Goods Vehicle", -1), vt_name_dict.get("Truck", -1)
# # #     if "TW" in s or "2W" in s:
# # #         return "TW", "Two Wheeler", subprod_dict.get("Two Wheeler", -1), -1
# # #     if "PRIVATE CAR" in s or re.search(r"\bPC\b", s):
# # #         return "PC", "Private Car", subprod_dict.get("Private Car", -1), vt_name_dict.get("Private Car", -1)
# # #     if "GCV" in s or "GVW" in s or "PCV" in s or "CV" in s:
# # #         return "CV", "Goods Vehicle", subprod_dict.get("Goods Vehicle", -1), vt_name_dict.get("Truck", -1)
# # #     return "CV", "Goods Vehicle", subprod_dict.get("Goods Vehicle", -1), vt_name_dict.get("Truck", -1)

# # # def infer_ncb_flag(text):
# # #     s = str(text).upper()
# # #     if "WITHOUT NCB" in s or "W/O NCB" in s or "NON NCB" in s:
# # #         return -1
# # #     if "WITH NCB" in s:
# # #         return 1
# # #     if "NCB" in s:
# # #         return 1
# # #     return -1

# # # def parse_remark_heuristic(remark_text):
# # #     s = keep_included_text(remark_text)
# # #     su = s.upper()
# # #     segment_hint = ""
# # #     if "GCV" in su:
# # #         segment_hint = "GCV"
# # #     elif "PCV" in su:
# # #         segment_hint = "PCV"
# # #     elif "TW" in su or "2W" in su:
# # #         segment_hint = "TW"
# # #     elif "PC" in su or "PRIVATE CAR" in su:
# # #         segment_hint = "PC"

# # #     make_name = ""
# # #     model_name = ""
# # #     make_match = re.search(r"MAKE\s*[:\-]\s*([A-Z0-9 &]+)", su)
# # #     if not make_match:
# # #         make_match = re.search(r"([A-Z0-9 &\-.]+?)\s+MAKE\s+ONLY", su)
# # #     model_match = re.search(r"MODEL\s*[:\-]\s*([A-Z0-9 &]+)", su)
# # #     if make_match:
# # #         make_name = make_match.group(1).strip().title()
# # #     if model_match:
# # #         model_name = model_match.group(1).strip().title()

# # #     return {
# # #         "segment_hint": segment_hint,
# # #         "vehicle_make": make_name,
# # #         "vehicle_model": model_name,
# # #         "vehicle_variant": "",
# # #         "is_with_ncb": infer_ncb_flag(su),
# # #     }

# # # def resolve_make_model_from_mmv(company_code, extracted_make, extracted_model, remark_text):
# # #     code = normalize_lookup_text(company_code)
# # #     make_name = str(extracted_make).strip()
# # #     model_name = str(extracted_model).strip()
# # #     make_norm = normalize_lookup_text(make_name)
# # #     model_norm = normalize_lookup_text(model_name)
# # #     make_id = -1
# # #     model_id = -1

# # #     if make_norm and (code, make_norm) in mmv_make_index:
# # #         make_id, canonical_make = mmv_make_index[(code, make_norm)]
# # #         if not make_name:
# # #             make_name = canonical_make
# # #     else:
# # #         if make_norm:
# # #             fuzzy = []
# # #             make_words = set(make_norm.split())
# # #             for cand_norm, cand_name, cand_id in mmv_company_make_names.get(code, []):
# # #                 cand_words = set(cand_norm.split())
# # #                 if make_norm in cand_norm or cand_norm in make_norm or (make_words and make_words.issubset(cand_words)):
# # #                     fuzzy.append((cand_norm, cand_name, cand_id))
# # #             if fuzzy:
# # #                 fuzzy.sort(key=lambda x: len(x[0]))
# # #                 make_norm, make_name, make_id = fuzzy[0]

# # #         remark_norm = normalize_lookup_text(keep_included_text(remark_text))
# # #         if make_id == -1:
# # #             for cand_norm, cand_name, cand_id in mmv_company_make_names.get(code, []):
# # #                 if re.search(rf"\b{re.escape(cand_norm)}\b", remark_norm):
# # #                     make_id = cand_id
# # #                     make_name = cand_name
# # #                     make_norm = cand_norm
# # #                     break
# # #         if make_id == -1:
# # #             mk = re.search(r"\b([A-Z0-9]+)\s+MAKE\s+ONLY\b", remark_norm)
# # #             if mk:
# # #                 token = mk.group(1)
# # #                 fuzzy = []
# # #                 for cand_norm, cand_name, cand_id in mmv_company_make_names.get(code, []):
# # #                     if token in cand_norm.split():
# # #                         fuzzy.append((cand_norm, cand_name, cand_id))
# # #                 if fuzzy:
# # #                     fuzzy.sort(key=lambda x: len(x[0]))
# # #                     make_norm, make_name, make_id = fuzzy[0]

# # #     included_norm = normalize_lookup_text(keep_included_text(remark_text))
# # #     if make_id != -1 and not model_norm:
# # #         for cand_model_norm, cand_model_name, cand_model_id in sorted(
# # #             list(mmv_models_by_make.get((code, make_norm), set())),
# # #             key=lambda x: len(x[0]),
# # #             reverse=True
# # #         ):
# # #             if re.search(rf"\b{re.escape(cand_model_norm)}\b", included_norm):
# # #                 model_norm = cand_model_norm
# # #                 model_name = cand_model_name
# # #                 model_id = cand_model_id
# # #                 break

# # #     if make_id != -1 and model_norm and (code, make_norm, model_norm) in mmv_model_index:
# # #         model_id, canonical_model = mmv_model_index[(code, make_norm, model_norm)]
# # #         if not model_name:
# # #             model_name = canonical_model
# # #     elif make_id == -1 and model_norm:
# # #         model_hits = mmv_company_model_names.get(code, {}).get(model_norm, [])
# # #         if len(model_hits) == 1:
# # #             model_id, canonical_model, make_id, canonical_make = model_hits[0]
# # #             model_name = canonical_model
# # #             make_name = canonical_make

# # #     return make_id, make_name, model_id, model_name

# # # def parse_remark_with_openai(remark_text, company_name, segment_text, policy_type):
# # #     api_key = os.getenv("OPENAI_API_KEY", "").strip()
# # #     if not api_key:
# # #         return parse_remark_heuristic(remark_text)

# # #     model = os.getenv("OPENAI_MODEL", "gpt-4.1-mini")
# # #     included_text = keep_included_text(remark_text)
# # #     prompt = (
# # #         "Extract structured insurance filters from this remark.\n"
# # #         "Return JSON only with keys: segment_hint, vehicle_make, vehicle_model, vehicle_variant, is_with_ncb.\n"
# # #         "Important: ignore rejected/declined/except/excluded parts. Return only INCLUDED make/model.\n"
# # #         "Example: 'tata altroz but eeco reject' => make='TATA', model='ALTROZ'.\n"
# # #         "segment_hint must be one of: GCV, PCV, TW, PC, CV, or empty string.\n"
# # #         "is_with_ncb must be 1 if NCB considered. If WITHOUT NCB/NON NCB or absent, return -1.\n"
# # #         f"company_name: {company_name}\n"
# # #         f"segment_text: {segment_text}\n"
# # #         f"policy_type: {policy_type}\n"
# # #         f"remark_original: {remark_text}\n"
# # #         f"remark_included_only: {included_text}\n"
# # #     )

# # #     body = {
# # #         "model": model,
# # #         "messages": [
# # #             {"role": "system", "content": "You are a strict JSON extractor."},
# # #             {"role": "user", "content": prompt}
# # #         ],
# # #         "response_format": {"type": "json_object"},
# # #         "temperature": 0
# # #     }

# # #     req = urllib.request.Request(
# # #         "https://api.openai.com/v1/chat/completions",
# # #         data=json.dumps(body).encode("utf-8"),
# # #         headers={
# # #             "Content-Type": "application/json",
# # #             "Authorization": f"Bearer {api_key}"
# # #         },
# # #         method="POST"
# # #     )

# # #     try:
# # #         with urllib.request.urlopen(req, timeout=30) as resp:
# # #             payload = json.loads(resp.read().decode("utf-8"))
# # #         content = payload["choices"][0]["message"]["content"]
# # #         parsed = json.loads(content)
# # #         return {
# # #             "segment_hint": str(parsed.get("segment_hint", "")).strip().upper(),
# # #             "vehicle_make": str(parsed.get("vehicle_make", "")).strip(),
# # #             "vehicle_model": str(parsed.get("vehicle_model", "")).strip(),
# # #             "vehicle_variant": str(parsed.get("vehicle_variant", "")).strip(),
# # #             "is_with_ncb": int(parsed.get("is_with_ncb", -1)),
# # #         }
# # #     except (urllib.error.URLError, urllib.error.HTTPError, KeyError, ValueError, TypeError):
# # #         return parse_remark_heuristic(remark_text)

# # # def build_lob_name(company_name, segment_hint, segment_text, policy_type, remark_text):
# # #     policy_label = policy_map.get(policy_type, "TP Only")
# # #     clean_remark = keep_included_text(remark_text)
# # #     parts = [
# # #         str(company_name).strip(),
# # #         str(segment_hint).strip() if str(segment_hint).strip() else str(segment_text).strip(),
# # #         str(policy_label).strip(),
# # #         str(clean_remark).strip(),
# # #     ]
# # #     return " ".join([p for p in parts if p])

# # # def get_cv_vehicle_info_from_tonnage(tonnage):
# # #     """
# # #     Maps Tonnage column value to vehicle type / subproduct / wheels
# # #     """
# # #     s = str(tonnage).strip().upper()
# # #     if 'PCV 3W' in s:
# # #         return vt_name_dict.get('Auto rikshaw', -1), 'Passenger Vehicle', subprod_dict.get('Passenger Vehicle', -1), 1, 0, 3
# # #     if 'GCV 3W' in s:
# # #         return vt_name_dict.get('GCV 3W Delivery Van', -1), 'Goods Vehicle', subprod_dict.get('Goods Vehicle', -1), 1, 0, 3
# # #     if 'MISD' in s or 'TRACTOR' in s:
# # #         return vt_name_dict.get('Agriculture Tractor', -1), 'Miscellaneous Vehicle', subprod_dict.get('Miscellaneous Vehicle', -1), -1, -1, -1
# # #     # GVW / tonnage based -> Truck
# # #     return vt_name_dict.get('Truck', -1), 'Goods Vehicle', subprod_dict.get('Goods Vehicle', -1), -1, -1, -1

# # # def get_cv_vehicle_info(original_segment):
# # #     """Used for CV TP — parses vehicle info from Original Segment text"""
# # #     s = str(original_segment).strip().upper()
# # #     if 'PCV 3W' in s:
# # #         return vt_name_dict.get('Auto rikshaw', -1), 'Passenger Vehicle', subprod_dict.get('Passenger Vehicle', -1), 1, 0, 3
# # #     if 'GCV 3W' in s:
# # #         return vt_name_dict.get('GCV 3W Delivery Van', -1), 'Goods Vehicle', subprod_dict.get('Goods Vehicle', -1), 1, 0, 3
# # #     return vt_name_dict.get('Truck', -1), 'Goods Vehicle', subprod_dict.get('Goods Vehicle', -1), -1, -1, -1

# # # def parse_pc_segment(original_segment):
# # #     s = str(original_segment).strip().upper()
# # #     if 'NCB / NON NCB' in s:
# # #         ncb_flags = [1, 0]
# # #     elif 'NON NCB' in s:
# # #         ncb_flags = [0]
# # #     elif 'NCB' in s:
# # #         ncb_flags = [1]
# # #     else:
# # #         ncb_flags = [-1]

# # #     if 'PETROL' in s:
# # #         fuels = [(fuel_dict['PETROL'], 'PETROL')]
# # #     elif 'DIESEL' in s and 'OTHERS' in s:
# # #         fuels = OTHERS_FUELS
# # #     elif 'DIESEL' in s:
# # #         fuels = [(fuel_dict['DIESEL'], 'DIESEL')]
# # #     else:
# # #         fuels = [(-1, '')]

# # #     return [(fid, fname, ncb) for fid, fname in fuels for ncb in ncb_flags]

# # # # ─── Build a single output row dict ─────────────────────────────────────────
# # # def build_row(company_id, comp_code, seg_id, seg_name, sub_prod_id, sub_prod_name,
# # #               rto_group_id, rto_group_name,
# # #               payin_od_rate, payin_tp_rate, payout_od_rate, payout_tp_rate,
# # #               policy_type, fuel_type_id, fuel_type_name, is_with_ncb,
# # #               is_geared, is_cc_considered, from_cc, to_cc,
# # #               is_weightage_considered, from_weightage_kg, to_weightage_kg,
# # #               is_no_of_wheel, from_wheel, to_wheel, vt_id,
# # #               is_mmv_considered=-1,
# # #               lob_name='',
# # #               make_id=-1, vehicle_make='',
# # #               model_id=-1, vehicle_model='',
# # #               variant_id=-1, vehicle_variant='',
# # #               from_age_month=0, to_age_month=700):
# # #     return {
# # #         'id': 0,
# # #         'company_id': company_id,
# # #         'company_code': comp_code,
# # #         'segment_id': seg_id,
# # #         'segment': seg_name,
# # #         'subproduct_id': sub_prod_id,
# # #         'sub_product_name': sub_prod_name,
# # #         'lob_id': -1,
# # #         'lob_name': lob_name,
# # #         'business_type_id': -1,
# # #         'business_type': 'Not Considered',
# # #         'is_highend_lob': False,
# # #         'rto_group_id': rto_group_id,
# # #         'rto_group_name': rto_group_name,
# # #         'payin_od_rate': payin_od_rate,
# # #         'payin_tp_rate': payin_tp_rate,
# # #         'payout_od_rate': payout_od_rate,
# # #         'payout_tp_rate': payout_tp_rate,
# # #         'extra_tp_rate': 0,
# # #         'eff_from_date': '2026-01-01',
# # #         'eff_to_date': '2026-01-16',
# # #         'fuel_type_id': fuel_type_id,
# # #         'fuel_type': fuel_type_name,
# # #         'is_on_net': True if policy_type == 'COMP' else False,
# # #         'is_one_year_pay_on_newbusiness': -1,
# # #         'is_cpa_included': -1,
# # #         'is_geared_vehicle': is_geared,
# # #         'is_cc_considered': is_cc_considered,
# # #         'from_cc': from_cc,
# # #         'to_cc': to_cc,
# # #         'is_premium_considered': -1,
# # #         'from_premium': -1,
# # #         'to_premium': -1,
# # #         'is_mmv_considered': is_mmv_considered,
# # #         'make_id': make_id,
# # #         'vehicle_make': vehicle_make,
# # #         'model_id': model_id,
# # #         'vehicle_model': vehicle_model,
# # #         'variant_id': variant_id,
# # #         'vehicle_variant': vehicle_variant,
# # #         'is_seating_cap_consider': -1,
# # #         'from_seating_cap': -1,
# # #         'to_seating_cap': -1,
# # #         'is_no_of_wheel_consider': is_no_of_wheel,
# # #         'from_no_of_wheel': from_wheel,
# # #         'to_no_of_wheel': to_wheel,
# # #         'vehicle_type_id': vt_id,
# # #         'ppi_in': 0,
# # #         'ppi_out': 0,
# # #         'is_irda_tp_included': -1,
# # #         'is_longterm_renewal_pay': -1,
# # #         'is_weightage_considered': is_weightage_considered,
# # #         'from_weightage_kg': from_weightage_kg,
# # #         'to_weightage_kg': to_weightage_kg,
# # #         'is_nil_dep_considered': -1,
# # #         'is_organization_type': -1,
# # #         'from_age_month': from_age_month,
# # #         'to_age_month': to_age_month,
# # #         'is_with_ncb': is_with_ncb,
# # #         'is_idv_cap_consider': -1,
# # #         'from_idv': 0,
# # #         'to_idv': 0,
# # #         'is_breakin_consider': -1,
# # #         'is_active': True,
# # #     }

# # # # ─── Process a single file ───────────────────────────────────────────────────
# # # def process_file(input_file):
# # #     df = pd.read_excel(input_file)
# # #     print(f"\nProcessing {len(df)} records from: {os.path.basename(input_file)}")

# # #     is_shriram_format = (
# # #         comp_code == "SHRIRAM"
# # #         and "SEGMENT" in df.columns
# # #         and "LOCATION" in df.columns
# # #         and ("POLICY TYPE" in df.columns or "Policy Type" in df.columns)
# # #     )
# # #     if is_shriram_format:
# # #         print("  Detected Shriram input format")
# # #         rows = []
# # #         remark_cache = {}
# # #         for _, row in df.iterrows():
# # #             policy_type = normalize_policy_type(row.get("POLICY TYPE", row.get("Policy Type", "")))
# # #             seg_name_raw = policy_map.get(policy_type, "TP Only")
# # #             seg_id, seg_name = segment_dict.get(seg_name_raw, (3, "TP Only"))

# # #             segment_text = str(row.get("SEGMENT", row.get("Segment", ""))).strip()
# # #             _, sub_prod_name, sub_prod_id, vt_id = get_shriram_vehicle_info(segment_text)
# # #             rto_group_name = str(row.get("LOCATION", row.get("Location", ""))).strip()
# # #             remark_text = str(row.get("REMARK", row.get("Remark", ""))).strip()
# # #             company_name_in = str(row.get("COMPANY NAME", row.get("Company Name", comp_code))).strip()

# # #             payin_val = to_float(row.get("PAYIN", row.get("Payin", 0)))
# # #             payout_val = to_float(row.get("PAYOUT", row.get("Payout", row.get("Calculated Payout", 0))))
# # #             from_age, to_age = parse_shriram_age_band(row.get("AGE", row.get("Age", "")))

# # #             cache_key = (remark_text, company_name_in, segment_text, policy_type)
# # #             if cache_key not in remark_cache:
# # #                 remark_cache[cache_key] = parse_remark_with_openai(
# # #                     remark_text=remark_text,
# # #                     company_name=company_name_in,
# # #                     segment_text=segment_text,
# # #                     policy_type=policy_type
# # #                 )
# # #             remark_meta = remark_cache[cache_key]
# # #             make_id, vehicle_make, model_id, vehicle_model = resolve_make_model_from_mmv(
# # #                 company_code=comp_code,
# # #                 extracted_make=remark_meta.get("vehicle_make", ""),
# # #                 extracted_model=remark_meta.get("vehicle_model", ""),
# # #                 remark_text=remark_text
# # #             )
# # #             is_mmv_considered = 1 if (make_id != -1 or model_id != -1) else -1
# # #             lob_name = build_lob_name(
# # #                 company_name=company_name_in if company_name_in else comp_code,
# # #                 segment_hint=remark_meta.get("segment_hint", ""),
# # #                 segment_text=segment_text,
# # #                 policy_type=policy_type,
# # #                 remark_text=remark_text
# # #             )

# # #             if policy_type == "TP":
# # #                 payin_od_rate, payin_tp_rate = 0, payin_val
# # #                 payout_od_rate, payout_tp_rate = 0, payout_val
# # #             elif policy_type == "SAOD":
# # #                 payin_od_rate, payin_tp_rate = payin_val, 0
# # #                 payout_od_rate, payout_tp_rate = payout_val, 0
# # #             else:
# # #                 payin_od_rate, payin_tp_rate = payin_val, payin_val
# # #                 payout_od_rate, payout_tp_rate = payout_val, payout_val

# # #             rows.append(build_row(
# # #                 company_id, comp_code, seg_id, seg_name, sub_prod_id, sub_prod_name,
# # #                 0, rto_group_name,
# # #                 payin_od_rate, payin_tp_rate, payout_od_rate, payout_tp_rate,
# # #                 policy_type, -1, "", remark_meta.get("is_with_ncb", -1),
# # #                 -1, -1, 0, 99999,
# # #                 -1, 0, 99999, -1, -1, -1, vt_id,
# # #                 is_mmv_considered=is_mmv_considered,
# # #                 lob_name=lob_name,
# # #                 make_id=make_id,
# # #                 vehicle_make=vehicle_make if vehicle_make else remark_meta.get("vehicle_make", ""),
# # #                 model_id=model_id,
# # #                 vehicle_model=vehicle_model if vehicle_model else remark_meta.get("vehicle_model", ""),
# # #                 variant_id=-1,
# # #                 vehicle_variant=remark_meta.get("vehicle_variant", ""),
# # #                 from_age_month=from_age, to_age_month=to_age
# # #             ))

# # #         print(f"  Expanded to {len(rows)} output rows")
# # #         return pd.DataFrame(rows)

# # #     lob_values = df['LOB'].dropna().unique().tolist() if 'LOB' in df.columns else []
# # #     lob_upper  = [str(l).strip().upper() for l in lob_values]
# # #     detected   = 'TW' if any(l == 'TW' for l in lob_upper) else ('CV' if any(l == 'CV' for l in lob_upper) else 'PC')
# # #     print(f"  Detected LOB type: {detected}")

# # #     # CV COMP: forward-fill Tonnage column so age band rows carry the tonnage
# # #     has_tonnage  = 'Tonnage' in df.columns
# # #     has_age_band = 'Age Band' in df.columns
# # #     if has_tonnage:
# # #         df['Tonnage'] = df['Tonnage'].ffill()

# # #     # Detect if this is CV COMP (has Tonnage + Age Band)
# # #     is_cv_comp = detected == 'CV' and has_tonnage and has_age_band

# # #     # If file has no Policy Type column, ask user explicitly ONCE
# # #     if 'Policy Type' not in df.columns or df['Policy Type'].dropna().empty:
# # #         print("\n" + "="*80)
# # #         print("  Policy Type not found in file. Please select:")
# # #         print("  1. COMP")
# # #         print("  2. TP")
# # #         print("  3. SAOD")
# # #         pt_choice = input("Enter choice: ").strip()
# # #         pt_map = {'1': 'COMP', '2': 'TP', '3': 'SAOD'}
# # #         default_policy_type = normalize_policy_type(pt_map.get(pt_choice, 'TP'))
# # #         print(f"✓ Using Policy Type: {default_policy_type}")
# # #     else:
# # #         default_policy_type = None

# # #     rows = []
# # #     for idx, row in df.iterrows():

# # #         # Policy Type
# # #         policy_raw = str(row.get('Policy Type', '')).strip()
# # #         if policy_raw in ('nan', ''):
# # #             policy_type = default_policy_type if default_policy_type else 'TP'
# # #         else:
# # #             policy_type = normalize_policy_type(policy_raw)
# # #         seg_name_raw     = policy_map.get(policy_type, 'TP Only')
# # #         seg_id, seg_name = segment_dict.get(seg_name_raw, (3, 'TP Only'))

# # #         lob              = str(row.get('LOB', 'TW')).strip().upper()
# # #         original_segment = str(row.get('Original Segment', '')).strip()
# # #         geo_new          = str(row.get('Geo Location', row.get('Geo New', ''))).strip()
# # #         rto_group_id     = 0
# # #         rto_group_name   = geo_new

# # #         # Payin column
# # #         payin_col  = 'Payin (OD Premium)' if 'Payin (OD Premium)' in df.columns else 'Payin'
# # #         payin_val  = to_float(row.get(payin_col, 0))
# # #         payout_val = to_float(row.get('Calculated Payout', 0))

# # #         if policy_type == 'TP':
# # #             payin_od_rate, payin_tp_rate   = 0, payin_val
# # #             payout_od_rate, payout_tp_rate = 0, payout_val
# # #         elif policy_type == 'SAOD':
# # #             payin_od_rate, payin_tp_rate   = payin_val, 0
# # #             payout_od_rate, payout_tp_rate = payout_val, 0
# # #         else:  # COMP
# # #             payin_od_rate, payin_tp_rate   = payin_val, payin_val
# # #             payout_od_rate, payout_tp_rate = payout_val, payout_val

# # #         # ── TW ───────────────────────────────────────────────────────────────
# # #         if lob == 'TW':
# # #             sub_prod_name = 'Two Wheeler'
# # #             sub_prod_id   = subprod_dict.get(sub_prod_name, -1)
# # #             tw_type       = str(row.get('TW Type', '')).strip()
# # #             is_geared     = 0 if 'scooter' in tw_type.lower() else 1
# # #             vt_id         = get_tw_vehicle_type_id(tw_type)
# # #             from_cc, to_cc, _ = parse_cc_band(row.get('CC Band', ''))
# # #             is_cc_considered  = get_is_cc_considered(original_segment)
# # #             rows.append(build_row(
# # #                 company_id, comp_code, seg_id, seg_name, sub_prod_id, sub_prod_name,
# # #                 rto_group_id, rto_group_name,
# # #                 payin_od_rate, payin_tp_rate, payout_od_rate, payout_tp_rate,
# # #                 policy_type, -1, '', -1,
# # #                 is_geared, is_cc_considered, from_cc, to_cc,
# # #                 -1, 0, 99999, -1, -1, -1, vt_id
# # #             ))

# # #         # ── CV COMP (has Tonnage + Age Band columns) ──────────────────────────
# # #         elif lob == 'CV' and is_cv_comp:
# # #             tonnage  = str(row.get('Tonnage', '')).strip()
# # #             age_band = str(row.get('Age Band', '')).strip()

# # #             # Vehicle info from Tonnage column
# # #             vt_id, sub_prod_name, sub_prod_id, is_no_of_wheel, from_wheel, to_wheel = get_cv_vehicle_info_from_tonnage(tonnage)

# # #             # Weight from Tonnage column
# # #             is_wt, from_wt, to_wt = parse_tonnage_col(tonnage)

# # #             # Age band -> months
# # #             from_age, to_age = parse_age_band(age_band)

# # #             rows.append(build_row(
# # #                 company_id, comp_code, seg_id, seg_name, sub_prod_id, sub_prod_name,
# # #                 rto_group_id, rto_group_name,
# # #                 payin_od_rate, payin_tp_rate, payout_od_rate, payout_tp_rate,
# # #                 policy_type, -1, '', -1,
# # #                 -1, -1, 0, 99999,
# # #                 is_wt, from_wt, to_wt, is_no_of_wheel, from_wheel, to_wheel, vt_id,
# # #                 from_age_month=from_age, to_age_month=to_age
# # #             ))

# # #         # ── CV TP (Original Segment based) ────────────────────────────────────
# # #         elif lob == 'CV':
# # #             vt_id, sub_prod_name, sub_prod_id, is_no_of_wheel, from_wheel, to_wheel = get_cv_vehicle_info(original_segment)
# # #             is_wt, from_wt, to_wt = parse_weight(original_segment)
# # #             rows.append(build_row(
# # #                 company_id, comp_code, seg_id, seg_name, sub_prod_id, sub_prod_name,
# # #                 rto_group_id, rto_group_name,
# # #                 payin_od_rate, payin_tp_rate, payout_od_rate, payout_tp_rate,
# # #                 policy_type, -1, '', -1,
# # #                 -1, -1, 0, 99999,
# # #                 is_wt, from_wt, to_wt, is_no_of_wheel, from_wheel, to_wheel, vt_id
# # #             ))

# # #         # ── PC ───────────────────────────────────────────────────────────────
# # #         elif lob in ('PC', 'PVT CAR'):
# # #             sub_prod_name = 'Private Car'
# # #             sub_prod_id   = subprod_dict.get(sub_prod_name, -1)
# # #             vt_id         = vt_name_dict.get('Private Car', -1)

# # #             direct_fuel = str(row.get('Fuel Type', '')).strip()
# # #             direct_cc   = str(row.get('CC Band', '')).strip()

# # #             if direct_fuel and direct_fuel.lower() != 'nan':
# # #                 # PC TP — Fuel Type and CC Band columns exist directly
# # #                 fuel_id   = fuel_dict.get(direct_fuel.upper(), -1)
# # #                 fuel_name = direct_fuel.upper() if fuel_id != -1 else ''
# # #                 from_cc, to_cc, is_cc = parse_cc_band(direct_cc)
# # #                 rows.append(build_row(
# # #                     company_id, comp_code, seg_id, seg_name, sub_prod_id, sub_prod_name,
# # #                     rto_group_id, rto_group_name,
# # #                     payin_od_rate, payin_tp_rate, payout_od_rate, payout_tp_rate,
# # #                     policy_type, fuel_id, fuel_name, -1,
# # #                     -1, is_cc, from_cc, to_cc,
# # #                     -1, 0, 99999, -1, -1, -1, vt_id
# # #                 ))
# # #             else:
# # #                 # PC COMP — parse fuel + NCB from Original Segment, expand rows
# # #                 expansions = parse_pc_segment(original_segment)
# # #                 for fuel_id, fuel_name, is_with_ncb in expansions:
# # #                     rows.append(build_row(
# # #                         company_id, comp_code, seg_id, seg_name, sub_prod_id, sub_prod_name,
# # #                         rto_group_id, rto_group_name,
# # #                         payin_od_rate, payin_tp_rate, payout_od_rate, payout_tp_rate,
# # #                         policy_type, fuel_id, fuel_name, is_with_ncb,
# # #                         -1, -1, 0, 99999,
# # #                         -1, 0, 99999, -1, -1, -1, vt_id
# # #                     ))

# # #         # ── Fallback ──
# # #         else:
# # #             sub_prod_name = lob_map.get(lob, lob)
# # #             sub_prod_id   = subprod_dict.get(sub_prod_name, -1)
# # #             rows.append(build_row(
# # #                 company_id, comp_code, seg_id, seg_name, sub_prod_id, sub_prod_name,
# # #                 rto_group_id, rto_group_name,
# # #                 payin_od_rate, payin_tp_rate, payout_od_rate, payout_tp_rate,
# # #                 policy_type, -1, '', -1,
# # #                 -1, -1, 0, 99999,
# # #                 -1, 0, 99999, -1, -1, -1, -1
# # #             ))

# # #     print(f"  Expanded to {len(rows)} output rows")
# # #     return pd.DataFrame(rows)

# # # # ─── Column order ────────────────────────────────────────────────────────────
# # # col_order = [
# # #     'id', 'company_id', 'company_code', 'segment_id', 'segment', 'subproduct_id', 'sub_product_name',
# # #     'lob_id', 'lob_name', 'business_type_id', 'business_type', 'is_highend_lob',
# # #     'rto_group_id', 'rto_group_name', 'payin_od_rate', 'payin_tp_rate', 'payout_od_rate', 'payout_tp_rate',
# # #     'extra_tp_rate', 'eff_from_date', 'eff_to_date', 'fuel_type_id', 'fuel_type',
# # #     'is_on_net', 'is_one_year_pay_on_newbusiness', 'is_cpa_included', 'is_geared_vehicle',
# # #     'is_cc_considered', 'from_cc', 'to_cc', 'is_premium_considered', 'from_premium', 'to_premium',
# # #     'is_mmv_considered', 'make_id', 'vehicle_make', 'model_id', 'vehicle_model', 'variant_id', 'vehicle_variant',
# # #     'is_seating_cap_consider', 'from_seating_cap', 'to_seating_cap',
# # #     'is_no_of_wheel_consider', 'from_no_of_wheel', 'to_no_of_wheel',
# # #     'vehicle_type_id', 'ppi_in', 'ppi_out',
# # #     'is_irda_tp_included', 'is_longterm_renewal_pay', 'is_weightage_considered',
# # #     'from_weightage_kg', 'to_weightage_kg', 'is_nil_dep_considered', 'is_organization_type',
# # #     'from_age_month', 'to_age_month', 'is_with_ncb', 'is_idv_cap_consider',
# # #     'from_idv', 'to_idv', 'is_breakin_consider', 'is_active'
# # # ]

# # # # ─── Main loop ───────────────────────────────────────────────────────────────
# # # input_file = input("\nEnter input Excel file path  : ").strip().strip('"')

# # # while True:
# # #     try:
# # #         out_df = process_file(input_file)
# # #         out_df = out_df[col_order]

# # #         if os.path.exists(output_path):
# # #             existing_df = pd.read_excel(output_path)
# # #             out_df = pd.concat([existing_df, out_df], ignore_index=True)
# # #             print(f"\n✓ Appended to existing file!")

# # #         out_df.to_excel(output_path, index=False)
# # #         print(f"✓ Output saved : {output_path}")
# # #         print(f"  Total records: {len(out_df)}")

# # #     except Exception as e:
# # #         print(f"\nERROR processing file: {e}")

# # #     print("\n" + "="*80)
# # #     print("  1. Yes - Add more files")
# # #     print("  2. No  - Exit")
# # #     choice = input("Do you want to add more files? Enter choice: ").strip()

# # #     if choice == '2':
# # #         print("\n✓ Done! Exiting.")
# # #         break
# # #     else:
# # #         input_file = input("\nEnter next input Excel file path : ").strip().strip('"')


# # # import pandas as pd
# # # import json
# # # import re
# # # import sys
# # # import os
# # # import urllib.request
# # # import urllib.error

# # # # ─── Load JSON reference data ────────────────────────────────────────────────
# # # def load_json(path):
# # #     with open(path) as f:
# # #         return json.load(f)

# # # def load_dotenv_if_exists(dotenv_path):
# # #     if not os.path.exists(dotenv_path):
# # #         return
# # #     with open(dotenv_path, encoding="utf-8") as f:
# # #         for line in f:
# # #             s = line.strip()
# # #             if not s or s.startswith("#") or "=" not in s:
# # #                 continue
# # #             k, v = s.split("=", 1)
# # #             k = k.strip()
# # #             v = v.strip().strip('"').strip("'")
# # #             if k and k not in os.environ:
# # #                 os.environ[k] = v

# # # def normalize_lookup_text(v):
# # #     s = str(v).upper()
# # #     s = re.sub(r'[^A-Z0-9]+', ' ', s)
# # #     return re.sub(r'\s+', ' ', s).strip()

# # # # ─── Ask for paths in CMD ─────────────────────────────────────────────────────
# # # print("\n" + "="*80)
# # # print("  SHRIRAM GENERAL INSURANCE — PayinConfig Generator")
# # # print("="*80)
# # # BASE          = input("Enter JSON files folder path   : ").strip().strip('"')
# # # output_folder = input("Enter output folder path       : ").strip().strip('"')
# # # load_dotenv_if_exists(os.path.join(BASE, ".env"))
# # # load_dotenv_if_exists(".env")

# # # companies     = load_json(os.path.join(BASE, 'company_master.json'))
# # # subproducts   = load_json(os.path.join(BASE, 'subproduct.json'))
# # # vehicle_types = load_json(os.path.join(BASE, 'vehicle_type.json'))
# # # rto_list      = load_json(os.path.join(BASE, 'rto_id_name.json'))
# # # fuel_list     = load_json(os.path.join(BASE, 'fuel.json'))

# # # # cv_mmv_master.json — optional, for make/model resolution
# # # for _p in [os.path.join(BASE, 'cv_mmv_master.json'), 'cv_mmv_master.json']:
# # #     if os.path.exists(_p):
# # #         cv_mmv_list = load_json(_p)
# # #         break
# # # else:
# # #     cv_mmv_list = []

# # # # ─── Build lookup dicts ───────────────────────────────────────────────────────
# # # company_dict  = {c['company_id']: c for c in companies}
# # # segment_dict  = {
# # #     'Comprehensive': (1, 'Comprehensive'),
# # #     'SAOD':          (2, 'SAOD'),
# # #     'TP Only':       (3, 'TP Only'),
# # # }
# # # policy_map    = {'COMP': 'Comprehensive', 'TP': 'TP Only', 'SAOD': 'SAOD'}
# # # subprod_dict  = {s['sub_product_name']: s['sub_product_id'] for s in subproducts}
# # # vt_name_dict  = {vt['vehicle_type']: vt['id'] for vt in vehicle_types}
# # # rto_dict      = {r['name']: r['id'] for r in rto_list}
# # # fuel_dict     = {f['fuel_type_name'].upper(): f['fuel_type_id'] for f in fuel_list}

# # # # MMV indexes
# # # mmv_make_index          = {}
# # # mmv_model_index         = {}
# # # mmv_company_make_names  = {}
# # # mmv_company_model_names = {}
# # # mmv_models_by_make      = {}

# # # for m in cv_mmv_list:
# # #     code       = normalize_lookup_text(m.get("company_code", ""))
# # #     make_name  = str(m.get("vehicle_make", "")).strip()
# # #     model_name = str(m.get("vehicle_model", "")).strip()
# # #     make_norm  = normalize_lookup_text(make_name)
# # #     model_norm = normalize_lookup_text(model_name)
# # #     make_id    = int(m.get("MakeID",  -1)) if str(m.get("MakeID",  "")).strip() not in ("", "nan", "None") else -1
# # #     model_id   = int(m.get("ModelID", -1)) if str(m.get("ModelID", "")).strip() not in ("", "nan", "None") else -1
# # #     if code and make_norm and make_id != -1:
# # #         mmv_make_index[(code, make_norm)] = (make_id, make_name)
# # #         mmv_company_make_names.setdefault(code, set()).add((make_norm, make_name, make_id))
# # #     if code and make_norm and model_norm and model_id != -1:
# # #         mmv_model_index[(code, make_norm, model_norm)] = (model_id, model_name)
# # #         mmv_company_model_names.setdefault(code, {}).setdefault(model_norm, []).append(
# # #             (model_id, model_name, make_id, make_name))
# # #         mmv_models_by_make.setdefault((code, make_norm), set()).add((model_norm, model_name, model_id))

# # # for k in list(mmv_company_make_names.keys()):
# # #     mmv_company_make_names[k] = sorted(
# # #         list(mmv_company_make_names[k]), key=lambda x: len(x[0]), reverse=True)

# # # # ─── Company selection ────────────────────────────────────────────────────────
# # # print("\n" + "="*80)
# # # print("AVAILABLE COMPANIES")
# # # print("="*80)
# # # for c in companies:
# # #     print(f"  ID: {c['company_id']:3d} | Code: {c['company_code']:20s} | {c['company_name']}")
# # # print("="*80)

# # # company_id_input = input("\nEnter company_id from the list above: ").strip()
# # # try:
# # #     company_id = int(company_id_input)
# # #     company    = company_dict[company_id]
# # #     comp_code  = company['company_code']
# # #     print(f"\n✓ Selected: {company['company_name']} (ID: {company_id}, Code: {comp_code})")
# # # except (ValueError, KeyError):
# # #     print(f"ERROR: Invalid company_id '{company_id_input}'")
# # #     sys.exit(1)

# # # output_path = os.path.join(output_folder, f'{comp_code}-Payin-Config.xlsx')

# # # # ─── Helpers ──────────────────────────────────────────────────────────────────
# # # def to_float(v):
# # #     if isinstance(v, str):
# # #         v = v.strip().replace('%', '')
# # #     try:
# # #         return float(v)
# # #     except:
# # #         return 0.0

# # # def normalize_policy_type(v):
# # #     s = str(v).strip().upper()
# # #     if s in ("COMP", "COMPREHENSIVE"):        return "COMP"
# # #     if s in ("TP", "TP ONLY", "THIRD PARTY"): return "TP"
# # #     if s in ("SAOD", "OD", "OWN DAMAGE"):     return "SAOD"
# # #     return "TP"

# # # def parse_cc_band(cc_band):
# # #     if not cc_band or str(cc_band).strip() == '' or str(cc_band).strip().lower() == 'nan':
# # #         return 0, 99999, -1
# # #     s = str(cc_band).strip().upper().replace('CC', '').strip()
# # #     m = re.match(r'^<\s*(\d+)$', s)
# # #     if m: return 0, int(m.group(1)) - 1, 1
# # #     m = re.match(r'^>\s*(\d+)$', s)
# # #     if m: return int(m.group(1)) + 1, 99999, 1
# # #     m = re.match(r'^>\s*(\d+)\s*[-–]\s*(\d+)$', s)
# # #     if m: return int(m.group(1)) + 1, int(m.group(2)), 1
# # #     m = re.match(r'^>=\s*(\d+)\s*[-–]\s*(\d+)$', s)
# # #     if m: return int(m.group(1)), int(m.group(2)), 1
# # #     m = re.match(r'^(\d+)\s*[-–]\s*(\d+)$', s)
# # #     if m: return int(m.group(1)), int(m.group(2)), 1
# # #     m = re.match(r'^(\d+)$', s)
# # #     if m: return int(m.group(1)), int(m.group(1)), 1
# # #     return 0, 99999, -1

# # # def parse_age(age_text):
# # #     """
# # #     Parse age from Shriram age column.
# # #     Supports: 'Upto X years', '>X - Y years', '>X years', blank -> (0, 700)
# # #     Returns (from_age_month, to_age_month)
# # #     """
# # #     s = str(age_text).strip()
# # #     if s == "" or s.lower() in ("nan", "none"):
# # #         return 0, 700
# # #     su = s.upper()

# # #     # Upto X Years  / Upto X Months
# # #     m = re.search(r'UPTO\s+(\d+)\s*(YEAR|YR|MONTH|MTH|MO)', su)
# # #     if m:
# # #         n, unit = int(m.group(1)), m.group(2)
# # #         return 0, n * 12 if unit.startswith('Y') else n

# # #     # >X - Y Years
# # #     m = re.match(r'>\s*(\d+)\s*[-–]\s*(\d+)(\+?)\s*[Yy]ear', s)
# # #     if m:
# # #         lo = int(m.group(1)) * 12 + 1
# # #         hi = 700 if m.group(3) == '+' else int(m.group(2)) * 12
# # #         return lo, hi

# # #     # >X+ Years  / >X Years
# # #     m = re.match(r'>\s*(\d+)\+?\s*[Yy]ear', s)
# # #     if m:
# # #         return int(m.group(1)) * 12 + 1, 700

# # #     # Plain number interpreted as years
# # #     m = re.match(r'^(\d+)$', s.strip())
# # #     if m:
# # #         return 0, int(m.group(1)) * 12

# # #     return 0, 700

# # # def parse_idv(remark_text):
# # #     """
# # #     Extract IDV cap from remark. Returns (is_idv_cap_consider, from_idv, to_idv)
# # #     e.g. 'IDV upto 10 lacs' -> (1, 0, 10)
# # #     """
# # #     su = str(remark_text).upper()
# # #     m = re.search(r'IDV\s+(?:UPTO|UP\s*TO|<=?)\s*([\d.]+)\s*(LAC|LAKH|L\b)', su)
# # #     if m:
# # #         val = float(m.group(1))
# # #         return 1, 0, val
# # #     m = re.search(r'IDV\s+([\d.]+)\s*[-–]\s*([\d.]+)\s*(LAC|LAKH)', su)
# # #     if m:
# # #         return 1, float(m.group(1)), float(m.group(2))
# # #     return -1, 0, 0

# # # def infer_ncb_flag(text):
# # #     s = str(text).upper()
# # #     if "WITHOUT NCB" in s or "W/O NCB" in s or "NON NCB" in s or "NON-NCB" in s:
# # #         return -1
# # #     if "NCB" in s:
# # #         return 1
# # #     return -1

# # # def infer_irda_tp(text):
# # #     s = str(text).upper()
# # #     if "IRDA TP" in s or "IRDA RATE" in s or "IRDA" in s:
# # #         return 1
# # #     return -1

# # # def keep_included_text(remark_text):
# # #     """Strip declined/excluded parts; keep only the included side."""
# # #     s = str(remark_text).strip()
# # #     if not s:
# # #         return ""
# # #     parts = re.split(r'[;,]', s)
# # #     kept = []
# # #     for p in parts:
# # #         up = p.upper()
# # #         if any(tok in up for tok in ("DECLIN", "REJECT")):
# # #             continue
# # #         for token in (" BUT ", " EXCEPT ", " EXCLUDE ", " OTHER THAN ", " NOT CONSIDER"):
# # #             idx = up.find(token)
# # #             if idx != -1:
# # #                 p = p[:idx]
# # #                 break
# # #         p = p.strip()
# # #         if p:
# # #             kept.append(p)
# # #     return " ".join([x for x in kept if x])

# # # def get_shriram_vehicle_info(segment_text):
# # #     """
# # #     Infer sub_product_name, sub_product_id, and vehicle_type_id from SEGMENT column.
# # #     Handles: TW, PC / PVT CAR / PRIVATE CAR, GCV, PCV, CV, All GVW etc.
# # #     """
# # #     s = str(segment_text).strip().upper()

# # #     # Two Wheeler patterns
# # #     if re.search(r'\bTW\b', s) or re.search(r'\b2W\b', s):
# # #         return "Two Wheeler", subprod_dict.get("Two Wheeler", -1), -1

# # #     # Private Car
# # #     if re.search(r'\bPRIVATE CAR\b', s) or re.search(r'\bPVT\s*CAR\b', s) or re.search(r'\bPC\b', s):
# # #         return "Private Car", subprod_dict.get("Private Car", -1), vt_name_dict.get("Private Car", -1)

# # #     # Passenger Vehicle — PCV, auto rickshaw, e-rickshaw, bus
# # #     if re.search(r'\bPCV\b', s) or re.search(r'\bPASSENGER\b', s):
# # #         return "Passenger Vehicle", subprod_dict.get("Passenger Vehicle", -1), vt_name_dict.get("Auto rikshaw", -1)

# # #     # GCV / CV / Goods
# # #     if (re.search(r'\bGCV\b', s) or re.search(r'\bGVW\b', s)
# # #             or "GOODS" in s or "ALL GVW" in s or re.search(r'\bCV\b', s)):
# # #         return "Goods Vehicle", subprod_dict.get("Goods Vehicle", -1), vt_name_dict.get("Truck", -1)

# # #     # Miscellaneous
# # #     if "TRACTOR" in s or "HARVESTER" in s or "MISC" in s:
# # #         return "Miscellaneous Vehicle", subprod_dict.get("Miscellaneous Vehicle", -1), vt_name_dict.get("Agriculture Tractor", -1)

# # #     # Default — treat as Goods Vehicle (CV)
# # #     return "Goods Vehicle", subprod_dict.get("Goods Vehicle", -1), vt_name_dict.get("Truck", -1)

# # # def get_tw_vehicle_type_id(tw_type_text):
# # #     s = str(tw_type_text).strip().lower()
# # #     if "scooter" in s:
# # #         return 17, 0   # TW Scooter, not geared
# # #     if "bike" in s:
# # #         return 18, 1   # TW Bike, geared
# # #     if "electric" in s:
# # #         return 20, -1  # TW Electric Bike
# # #     return -1, -1

# # # # ─── MMV Resolution ───────────────────────────────────────────────────────────
# # # def resolve_make_model_from_mmv(company_code, extracted_make, extracted_model, remark_text):
# # #     code       = normalize_lookup_text(company_code)
# # #     make_name  = str(extracted_make).strip()
# # #     model_name = str(extracted_model).strip()
# # #     make_norm  = normalize_lookup_text(make_name)
# # #     model_norm = normalize_lookup_text(model_name)
# # #     make_id    = -1
# # #     model_id   = -1

# # #     # Direct make lookup
# # #     if make_norm and (code, make_norm) in mmv_make_index:
# # #         make_id, make_name = mmv_make_index[(code, make_norm)]
# # #     else:
# # #         # Fuzzy make scan
# # #         if make_norm:
# # #             make_words = set(make_norm.split())
# # #             fuzzy = []
# # #             for cand_norm, cand_name, cand_id in mmv_company_make_names.get(code, []):
# # #                 cand_words = set(cand_norm.split())
# # #                 if make_norm in cand_norm or cand_norm in make_norm or make_words.issubset(cand_words):
# # #                     fuzzy.append((cand_norm, cand_name, cand_id))
# # #             if fuzzy:
# # #                 fuzzy.sort(key=lambda x: len(x[0]))
# # #                 make_norm, make_name, make_id = fuzzy[0]

# # #         # Scan included remark text for known makes
# # #         included_norm = normalize_lookup_text(keep_included_text(remark_text))
# # #         if make_id == -1:
# # #             for cand_norm, cand_name, cand_id in mmv_company_make_names.get(code, []):
# # #                 if re.search(rf"\b{re.escape(cand_norm)}\b", included_norm):
# # #                     make_id   = cand_id
# # #                     make_name = cand_name
# # #                     make_norm = cand_norm
# # #                     break

# # #     # Model lookup if make known
# # #     included_norm = normalize_lookup_text(keep_included_text(remark_text))
# # #     if make_id != -1 and not model_norm:
# # #         for cand_model_norm, cand_model_name, cand_model_id in sorted(
# # #             list(mmv_models_by_make.get((code, make_norm), set())),
# # #             key=lambda x: len(x[0]), reverse=True
# # #         ):
# # #             if re.search(rf"\b{re.escape(cand_model_norm)}\b", included_norm):
# # #                 model_norm = cand_model_norm
# # #                 model_name = cand_model_name
# # #                 model_id   = cand_model_id
# # #                 break

# # #     if make_id != -1 and model_norm and (code, make_norm, model_norm) in mmv_model_index:
# # #         model_id, canonical_model = mmv_model_index[(code, make_norm, model_norm)]
# # #         if not model_name:
# # #             model_name = canonical_model
# # #     elif make_id == -1 and model_norm:
# # #         model_hits = mmv_company_model_names.get(code, {}).get(model_norm, [])
# # #         if len(model_hits) == 1:
# # #             model_id, canonical_model, make_id, canonical_make = model_hits[0]
# # #             model_name = canonical_model
# # #             make_name  = canonical_make

# # #     return make_id, make_name, model_id, model_name

# # # # ─── OpenAI remark parser ─────────────────────────────────────────────────────
# # # def parse_remark_heuristic(remark_text, segment_text="", policy_type=""):
# # #     su = str(remark_text).upper()
# # #     return {
# # #         "segment_hint":    "",
# # #         "vehicle_make":    "",
# # #         "vehicle_model":   "",
# # #         "vehicle_variant": "",
# # #         "is_with_ncb":     infer_ncb_flag(su),
# # #         "is_irda_tp":      infer_irda_tp(su),
# # #     }

# # # def parse_remark_with_openai(remark_text, company_name, segment_text, policy_type):
# # #     api_key = os.getenv("OPENAI_API_KEY", "").strip()
# # #     if not api_key:
# # #         return parse_remark_heuristic(remark_text, segment_text, policy_type)

# # #     model         = os.getenv("OPENAI_MODEL", "gpt-4.1-mini")
# # #     included_text = keep_included_text(remark_text)

# # #     prompt = (
# # #         "Extract structured insurance filters from this remark for a motor insurance payin config system.\n"
# # #         "Return JSON ONLY with these exact keys:\n"
# # #         "  segment_hint       - one of: GCV, PCV, TW, PC, CV, or empty string\n"
# # #         "  vehicle_make       - make names INCLUDED (comma-separated if multiple, else empty string)\n"
# # #         "  vehicle_model      - model names INCLUDED (comma-separated if multiple, else empty string)\n"
# # #         "  vehicle_variant    - variant name or empty string\n"
# # #         "  is_with_ncb        - 1 if NCB cases included; -1 if WITHOUT NCB / NON NCB or not mentioned\n"
# # #         "  is_irda_tp         - 1 if IRDA TP rate is included/mentioned; -1 otherwise\n"
# # #         "Rules:\n"
# # #         "  - IGNORE rejected/declined/except/excluded parts, focus only on INCLUDED items.\n"
# # #         "  - If multiple makes, list them comma-separated (e.g. 'HONDA,HYUNDAI,KIA').\n"
# # #         "  - For model exceptions like 'EXCEPT ALTO/K10', do NOT include those as vehicle_model.\n"
# # #         "  - If OD only is mentioned, segment_hint can remain empty.\n"
# # #         f"company_name:           {company_name}\n"
# # #         f"segment_text:           {segment_text}\n"
# # #         f"policy_type:            {policy_type}\n"
# # #         f"remark_original:        {remark_text}\n"
# # #         f"remark_included_only:   {included_text}\n"
# # #     )

# # #     body = {
# # #         "model": model,
# # #         "messages": [
# # #             {"role": "system", "content": "You are a strict JSON extractor for insurance data."},
# # #             {"role": "user",   "content": prompt}
# # #         ],
# # #         "response_format": {"type": "json_object"},
# # #         "temperature": 0
# # #     }

# # #     req = urllib.request.Request(
# # #         "https://api.openai.com/v1/chat/completions",
# # #         data    = json.dumps(body).encode("utf-8"),
# # #         headers = {
# # #             "Content-Type":  "application/json",
# # #             "Authorization": f"Bearer {api_key}"
# # #         },
# # #         method  = "POST"
# # #     )

# # #     try:
# # #         with urllib.request.urlopen(req, timeout=30) as resp:
# # #             payload = json.loads(resp.read().decode("utf-8"))
# # #         content = payload["choices"][0]["message"]["content"]
# # #         parsed  = json.loads(content)
# # #         return {
# # #             "segment_hint":    str(parsed.get("segment_hint",    "")).strip().upper(),
# # #             "vehicle_make":    str(parsed.get("vehicle_make",    "")).strip(),
# # #             "vehicle_model":   str(parsed.get("vehicle_model",   "")).strip(),
# # #             "vehicle_variant": str(parsed.get("vehicle_variant", "")).strip(),
# # #             "is_with_ncb":     int(parsed.get("is_with_ncb",  -1)),
# # #             "is_irda_tp":      int(parsed.get("is_irda_tp",   -1)),
# # #         }
# # #     except Exception:
# # #         return parse_remark_heuristic(remark_text, segment_text, policy_type)

# # # # ─── Build one output row dict ────────────────────────────────────────────────
# # # def build_row(
# # #     company_id, comp_code,
# # #     seg_id, seg_name,
# # #     sub_prod_id, sub_prod_name,
# # #     rto_group_id, rto_group_name,
# # #     payin_od_rate, payin_tp_rate, payout_od_rate, payout_tp_rate,
# # #     policy_type,
# # #     fuel_type_id, fuel_type_name,
# # #     is_with_ncb,
# # #     is_geared,
# # #     is_cc_considered, from_cc, to_cc,
# # #     is_weightage_considered, from_weightage_kg, to_weightage_kg,
# # #     is_no_of_wheel, from_wheel, to_wheel,
# # #     vehicle_type_id,
# # #     # optional
# # #     is_mmv_considered = -1,
# # #     lob_name          = '',
# # #     make_id           = -1, vehicle_make    = '',
# # #     model_id          = -1, vehicle_model   = '',
# # #     variant_id        = -1, vehicle_variant = '',
# # #     from_age_month    = 0,  to_age_month    = 700,
# # #     is_idv_cap_consider = -1, from_idv = 0, to_idv = 0,
# # #     is_irda_tp_included = -1,
# # # ):
# # #     return {
# # #         'id':                           0,
# # #         'company_id':                   company_id,
# # #         'company_code':                 comp_code,
# # #         'segment_id':                   seg_id,
# # #         'segment':                      seg_name,
# # #         'subproduct_id':                sub_prod_id,
# # #         'sub_product_name':             sub_prod_name,
# # #         'lob_id':                       -1,
# # #         'lob_name':                     lob_name,
# # #         'business_type_id':             -1,
# # #         'business_type':                'Not Considered',
# # #         'is_highend_lob':               False,
# # #         'rto_group_id':                 rto_group_id,
# # #         'rto_group_name':               rto_group_name,
# # #         'payin_od_rate':                payin_od_rate,
# # #         'payin_tp_rate':                payin_tp_rate,
# # #         'payout_od_rate':               payout_od_rate,
# # #         'payout_tp_rate':               payout_tp_rate,
# # #         'extra_tp_rate':                0,
# # #         'eff_from_date':                '2026-01-01',
# # #         'eff_to_date':                  '2026-01-16',
# # #         'fuel_type_id':                 fuel_type_id,
# # #         'fuel_type':                    fuel_type_name,
# # #         'is_on_net':                    False,
# # #         'is_one_year_pay_on_newbusiness': -1,
# # #         'is_cpa_included':              -1,
# # #         'is_geared_vehicle':            is_geared,
# # #         'is_cc_considered':             is_cc_considered,
# # #         'from_cc':                      from_cc,
# # #         'to_cc':                        to_cc,
# # #         'is_premium_considered':        -1,
# # #         'from_premium':                 -1,
# # #         'to_premium':                   -1,
# # #         'is_mmv_considered':            is_mmv_considered,
# # #         'make_id':                      make_id,
# # #         'vehicle_make':                 vehicle_make,
# # #         'model_id':                     model_id,
# # #         'vehicle_model':                vehicle_model,
# # #         'variant_id':                   variant_id,
# # #         'vehicle_variant':              vehicle_variant,
# # #         'is_seating_cap_consider':      -1,
# # #         'from_seating_cap':             -1,
# # #         'to_seating_cap':               -1,
# # #         'is_no_of_wheel_consider':      is_no_of_wheel,
# # #         'from_no_of_wheel':             from_wheel,
# # #         'to_no_of_wheel':               to_wheel,
# # #         'vehicle_type_id':              vehicle_type_id,
# # #         'ppi_in':                       0,
# # #         'ppi_out':                      0,
# # #         'is_irda_tp_included':          is_irda_tp_included,
# # #         'is_longterm_renewal_pay':      -1,
# # #         'is_weightage_considered':      is_weightage_considered,
# # #         'from_weightage_kg':            from_weightage_kg,
# # #         'to_weightage_kg':              to_weightage_kg,
# # #         'is_nil_dep_considered':        -1,
# # #         'is_organization_type':         -1,
# # #         'from_age_month':               from_age_month,
# # #         'to_age_month':                 to_age_month,
# # #         'is_with_ncb':                  is_with_ncb,
# # #         'is_idv_cap_consider':          is_idv_cap_consider,
# # #         'from_idv':                     from_idv,
# # #         'to_idv':                       to_idv,
# # #         'is_breakin_consider':          -1,
# # #         'is_active':                    True,
# # #     }

# # # # ─── Core processing function ─────────────────────────────────────────────────
# # # def process_file(input_file):
# # #     df = pd.read_excel(input_file)
# # #     print(f"\n  Processing {len(df)} rows from: {os.path.basename(input_file)}")
# # #     print(f"  Columns found: {list(df.columns)}")

# # #     # Normalise column names (strip whitespace)
# # #     df.columns = [c.strip() for c in df.columns]

# # #     # Detect column aliases
# # #     def col(df, *names):
# # #         """Return first column name that exists in df, else None."""
# # #         for n in names:
# # #             if n in df.columns:
# # #                 return n
# # #         return None

# # #     col_segment     = col(df, 'SEGMENT', 'Segment', 'LOB')
# # #     col_policy_type = col(df, 'POLICY TYPE', 'Policy Type', 'POLICYTYPE')
# # #     col_location    = col(df, 'LOCATION', 'Location', 'GEO LOCATION', 'Geo Location')
# # #     col_payin       = col(df, 'PAYIN', 'Payin', 'PAYIN (OD)', 'Payin (OD Premium)')
# # #     col_payout      = col(df, 'PAYOUT', 'Payout', 'Calculated Payout', 'CALCULATED PAYOUT')
# # #     col_remark      = col(df, 'REMARK', 'Remark', 'REMARKS', 'Remarks')
# # #     col_age         = col(df, 'AGE', 'Age', 'AGE BAND', 'Age Band', 'AGE (YEARS)')
# # #     col_cc          = col(df, 'CC BAND', 'CC Band', 'CC', 'CC_BAND')
# # #     col_tw_type     = col(df, 'TW TYPE', 'TW Type', 'TW_TYPE')
# # #     col_company     = col(df, 'COMPANY NAME', 'Company Name', 'COMPANY', 'Company')

# # #     rows        = []
# # #     remark_cache = {}

# # #     for _, row in df.iterrows():

# # #         # ── Policy Type ───────────────────────────────────────────────────────
# # #         pt_raw      = str(row.get(col_policy_type, "TP") if col_policy_type else "TP").strip()
# # #         policy_type = normalize_policy_type(pt_raw)
# # #         seg_name_raw = policy_map.get(policy_type, "TP Only")
# # #         seg_id, seg_name = segment_dict.get(seg_name_raw, (3, "TP Only"))

# # #         # ── Payin / Payout ────────────────────────────────────────────────────
# # #         payin_val  = to_float(row.get(col_payin,  0) if col_payin  else 0)
# # #         payout_val = to_float(row.get(col_payout, 0) if col_payout else 0)

# # #         if policy_type == "TP":
# # #             payin_od_rate, payin_tp_rate   = 0, payin_val
# # #             payout_od_rate, payout_tp_rate = 0, payout_val
# # #         elif policy_type == "SAOD":
# # #             payin_od_rate, payin_tp_rate   = payin_val, 0
# # #             payout_od_rate, payout_tp_rate = payout_val, 0
# # #         else:  # COMP
# # #             payin_od_rate, payin_tp_rate   = payin_val, payin_val
# # #             payout_od_rate, payout_tp_rate = payout_val, payout_val

# # #         # ── Segment / Vehicle Info ────────────────────────────────────────────
# # #         segment_text = str(row.get(col_segment, "") if col_segment else "").strip()
# # #         sub_prod_name, sub_prod_id, vt_id = get_shriram_vehicle_info(segment_text)

# # #         # ── Location ──────────────────────────────────────────────────────────
# # #         rto_group_name = str(row.get(col_location, "") if col_location else "").strip()
# # #         rto_group_id   = 0  # Raw text; ID resolution done post-process if needed

# # #         # ── Remark ───────────────────────────────────────────────────────────
# # #         remark_text  = str(row.get(col_remark, "") if col_remark else "").strip()
# # #         company_name = str(row.get(col_company, comp_code) if col_company else comp_code).strip()

# # #         # ── Age ───────────────────────────────────────────────────────────────
# # #         age_raw = str(row.get(col_age, "") if col_age else "").strip()
# # #         from_age, to_age = parse_age(age_raw)

# # #         # ── CC Band ───────────────────────────────────────────────────────────
# # #         cc_raw = str(row.get(col_cc, "") if col_cc else "").strip()
# # #         from_cc, to_cc, is_cc = parse_cc_band(cc_raw)

# # #         # ── TW-specific geared / vehicle type ─────────────────────────────────
# # #         tw_type_raw = str(row.get(col_tw_type, "") if col_tw_type else "").strip()
# # #         if sub_prod_name == "Two Wheeler" and tw_type_raw:
# # #             vt_id, is_geared = get_tw_vehicle_type_id(tw_type_raw)
# # #         elif sub_prod_name == "Two Wheeler":
# # #             is_geared = -1
# # #         else:
# # #             is_geared = -1

# # #         # ── OpenAI remark parsing (with caching) ──────────────────────────────
# # #         cache_key = (remark_text, company_name, segment_text, policy_type)
# # #         if cache_key not in remark_cache:
# # #             remark_cache[cache_key] = parse_remark_with_openai(
# # #                 remark_text  = remark_text,
# # #                 company_name = company_name,
# # #                 segment_text = segment_text,
# # #                 policy_type  = policy_type
# # #             )
# # #         meta = remark_cache[cache_key]

# # #         is_with_ncb     = meta.get("is_with_ncb",  -1)
# # #         is_irda_tp      = meta.get("is_irda_tp",   -1)
# # #         raw_make_text   = meta.get("vehicle_make",  "")
# # #         raw_model_text  = meta.get("vehicle_model", "")

# # #         # ── Handle multiple makes (e.g. "HONDA,HYUNDAI,KIA") ──────────────────
# # #         # Each make → separate output row (same for models if needed)
# # #         make_list = [m.strip() for m in re.split(r'[,&/]+', raw_make_text) if m.strip()]
# # #         if not make_list:
# # #             make_list = [""]

# # #         for raw_make in make_list:
# # #             make_id, vehicle_make, model_id, vehicle_model = resolve_make_model_from_mmv(
# # #                 company_code   = comp_code,
# # #                 extracted_make = raw_make,
# # #                 extracted_model= raw_model_text,
# # #                 remark_text    = remark_text
# # #             )
# # #             is_mmv_considered = 1 if (make_id != -1 or model_id != -1 or raw_make or raw_model_text) else -1

# # #             # IDV
# # #             is_idv_cap, from_idv, to_idv = parse_idv(remark_text)

# # #             # LOB name string
# # #             lob_name = " ".join(filter(None, [
# # #                 company_name,
# # #                 segment_text,
# # #                 policy_type,
# # #                 keep_included_text(remark_text)
# # #             ]))

# # #             rows.append(build_row(
# # #                 company_id      = company_id,
# # #                 comp_code       = comp_code,
# # #                 seg_id          = seg_id,
# # #                 seg_name        = seg_name,
# # #                 sub_prod_id     = sub_prod_id,
# # #                 sub_prod_name   = sub_prod_name,
# # #                 rto_group_id    = rto_group_id,
# # #                 rto_group_name  = rto_group_name,
# # #                 payin_od_rate   = payin_od_rate,
# # #                 payin_tp_rate   = payin_tp_rate,
# # #                 payout_od_rate  = payout_od_rate,
# # #                 payout_tp_rate  = payout_tp_rate,
# # #                 policy_type     = policy_type,
# # #                 fuel_type_id    = -1,
# # #                 fuel_type_name  = '',
# # #                 is_with_ncb     = is_with_ncb,
# # #                 is_geared       = is_geared,
# # #                 is_cc_considered= is_cc,
# # #                 from_cc         = from_cc,
# # #                 to_cc           = to_cc,
# # #                 is_weightage_considered = -1,
# # #                 from_weightage_kg       = 0,
# # #                 to_weightage_kg         = 99999,
# # #                 is_no_of_wheel  = -1,
# # #                 from_wheel      = -1,
# # #                 to_wheel        = -1,
# # #                 vehicle_type_id = vt_id,
# # #                 is_mmv_considered    = is_mmv_considered,
# # #                 lob_name             = lob_name,
# # #                 make_id              = make_id,
# # #                 vehicle_make         = vehicle_make if vehicle_make else raw_make,
# # #                 model_id             = model_id,
# # #                 vehicle_model        = vehicle_model if vehicle_model else raw_model_text,
# # #                 variant_id           = -1,
# # #                 vehicle_variant      = meta.get("vehicle_variant", ""),
# # #                 from_age_month       = from_age,
# # #                 to_age_month         = to_age,
# # #                 is_idv_cap_consider  = is_idv_cap,
# # #                 from_idv             = from_idv,
# # #                 to_idv               = to_idv,
# # #                 is_irda_tp_included  = is_irda_tp,
# # #             ))

# # #     print(f"  Expanded to {len(rows)} output rows")
# # #     return pd.DataFrame(rows)

# # # # ─── Column order (matches master schema) ────────────────────────────────────
# # # col_order = [
# # #     'id', 'company_id', 'company_code', 'segment_id', 'segment',
# # #     'subproduct_id', 'sub_product_name',
# # #     'lob_id', 'lob_name', 'business_type_id', 'business_type', 'is_highend_lob',
# # #     'rto_group_id', 'rto_group_name',
# # #     'payin_od_rate', 'payin_tp_rate', 'payout_od_rate', 'payout_tp_rate',
# # #     'extra_tp_rate', 'eff_from_date', 'eff_to_date',
# # #     'fuel_type_id', 'fuel_type',
# # #     'is_on_net', 'is_one_year_pay_on_newbusiness', 'is_cpa_included',
# # #     'is_geared_vehicle', 'is_cc_considered', 'from_cc', 'to_cc',
# # #     'is_premium_considered', 'from_premium', 'to_premium',
# # #     'is_mmv_considered', 'make_id', 'vehicle_make', 'model_id', 'vehicle_model',
# # #     'variant_id', 'vehicle_variant',
# # #     'is_seating_cap_consider', 'from_seating_cap', 'to_seating_cap',
# # #     'is_no_of_wheel_consider', 'from_no_of_wheel', 'to_no_of_wheel',
# # #     'vehicle_type_id', 'ppi_in', 'ppi_out',
# # #     'is_irda_tp_included', 'is_longterm_renewal_pay', 'is_weightage_considered',
# # #     'from_weightage_kg', 'to_weightage_kg',
# # #     'is_nil_dep_considered', 'is_organization_type',
# # #     'from_age_month', 'to_age_month',
# # #     'is_with_ncb', 'is_idv_cap_consider', 'from_idv', 'to_idv',
# # #     'is_breakin_consider', 'is_active',
# # # ]

# # # # ─── Main loop ────────────────────────────────────────────────────────────────
# # # input_file = input("\nEnter Shriram input Excel file path: ").strip().strip('"')

# # # while True:
# # #     try:
# # #         out_df = process_file(input_file)
# # #         out_df = out_df[[c for c in col_order if c in out_df.columns]]

# # #         if os.path.exists(output_path):
# # #             existing_df = pd.read_excel(output_path)
# # #             out_df = pd.concat([existing_df, out_df], ignore_index=True)
# # #             print(f"\n✓ Appended to existing output file.")

# # #         out_df.to_excel(output_path, index=False)
# # #         print(f"✓ Output saved : {output_path}")
# # #         print(f"  Total records: {len(out_df)}")

# # #     except Exception as e:
# # #         import traceback
# # #         print(f"\nERROR: {e}")
# # #         traceback.print_exc()

# # #     print("\n" + "="*80)
# # #     print("  1. Add more Shriram files")
# # #     print("  2. Exit")
# # #     choice = input("Choice: ").strip()
# # #     if choice == "2":
# # #         print("\n✓ Done!")
# # #         break
# # #     input_file = input("\nEnter next Shriram input Excel file path: ").strip().strip('"')






# # import pandas as pd
# # import re, sys, os, json, time, urllib.request, urllib.error
# # from datetime import datetime

# # # =============================================================================
# # #  SHRIRAM GENERAL INSURANCE — PayinConfig Generator  v3
# # #  Reads master data from Masters_-_dec_2025.xlsx  +  JSON reference files
# # #  Run:  python SHRIRAM_PayinConfig_v3.py
# # # =============================================================================

# # # ── Logging ──────────────────────────────────────────────────────────────────
# # C = {
# #     "OK"   : "\033[32m", "WARN" : "\033[33m", "ERR"  : "\033[31m",
# #     "API"  : "\033[36m", "CACHE": "\033[90m", "INFO" : "",
# # }
# # _log_file = None

# # def log(level, msg):
# #     ts   = datetime.now().strftime("%H:%M:%S")
# #     line = f"[{ts}][{level}] {msg}"
# #     col  = C.get(level, "")
# #     print(f"{col}{line}\033[0m" if col else line)
# #     if _log_file:
# #         _log_file.write(line + "\n")
# #         _log_file.flush()

# # def bar(cur, tot, w=48):
# #     pct = cur / tot if tot else 0
# #     b   = "\u2588" * int(w * pct) + "\u2591" * (w - int(w * pct))
# #     print(f"\r  [{b}] {cur}/{tot} ({pct*100:.1f}%)", end="", flush=True)
# #     if cur == tot: print()

# # # ── Tiny helpers ──────────────────────────────────────────────────────────────
# # def nt(v):
# #     """Normalise: UPPER + collapse non-alnum to single space."""
# #     return re.sub(r'\s+', ' ', re.sub(r'[^A-Z0-9]+', ' ', str(v).upper())).strip()

# # def si(v, d=-1):
# #     try:   return int(float(str(v)))
# #     except: return d

# # def sf(v, d=0.0):
# #     if isinstance(v, str): v = v.strip().replace('%','')
# #     try:   return float(v)
# #     except: return d

# # def load_dotenv(p):
# #     if not os.path.exists(p): return
# #     for line in open(p, encoding="utf-8"):
# #         s = line.strip()
# #         if not s or s.startswith('#') or '=' not in s: continue
# #         k, v = s.split('=', 1)
# #         k = k.strip(); v = v.strip().strip('"').strip("'")
# #         if k and k not in os.environ: os.environ[k] = v

# # # =============================================================================
# # #  STARTUP
# # # =============================================================================
# # print("\n" + "="*75)
# # print("  SHRIRAM GENERAL INSURANCE — PayinConfig Generator  v3")
# # print("="*75)

# # MASTERS = input("Path to Masters_-_dec_2025.xlsx  : ").strip().strip('"')
# # JSON_DIR= input("Path to JSON files folder        : ").strip().strip('"')
# # OUT_DIR = input("Output folder path               : ").strip().strip('"')
# # os.makedirs(OUT_DIR, exist_ok=True)

# # _log_path = os.path.join(OUT_DIR, f"shriram_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
# # try:
# #     _log_file = open(_log_path, "w", encoding="utf-8")
# #     log("OK", f"Log → {_log_path}")
# # except Exception as e:
# #     print(f"[WARN] Cannot open log file: {e}")

# # load_dotenv(os.path.join(JSON_DIR, ".env"))
# # load_dotenv(".env")

# # # =============================================================================
# # #  LOAD MASTERS  (Excel sheets take priority; JSON files fill any gaps)
# # # =============================================================================
# # log("INFO", f"Loading master data from Excel: {MASTERS}")
# # t0  = time.time()
# # xl  = pd.ExcelFile(MASTERS)

# # def rsh(name):
# #     df = pd.read_excel(xl, sheet_name=name)
# #     df.columns = [str(c).strip() for c in df.columns]
# #     return df

# # df_company = rsh("Insurance Company Master")   # company_id, company_code, company_name
# # df_segment = rsh("Segment Master")            # id, segment
# # df_subprod = rsh("SubProduct Master")         # sub_product_id, sub_product_name (+product_name)
# # df_vtype   = rsh("Vehicle type")              # id, vehicle_type, no_of_wheels, sub_product_name
# # df_fuel_xl = rsh("fuel master")               # id, fuel_type
# # df_tw_mmv  = rsh("tw_pc_mmv master")          # MakeID/vehicle_make/ModelID/vehicle_model/VariantID/VariantDisplayName/CC/Fuel/SeatingCapacity/Is_Geared_Vehicle
# # df_cv_mmv  = rsh("cv_mmv master")             # company_code/MakeID/vehicle_make/ModelID/vehicle_model/VariantID/VariantName/Fuel/SeatingCapacity

# # log("OK", f"Excel masters loaded in {time.time()-t0:.1f}s")

# # # ── Company dict ──────────────────────────────────────────────────────────────
# # company_by_id = {si(r['company_id']): r for _, r in df_company.iterrows()}

# # # ── Segment  (id → name, name → id) ──────────────────────────────────────────
# # seg_name_to_id = {str(r['segment']).strip(): si(r['id']) for _, r in df_segment.iterrows()}
# # # Policy-type string → canonical segment name
# # PT_SEG = {
# #     'COMPREHENSIVE': 'Comprehensive', 'COMP': 'Comprehensive',
# #     'TP': 'TP Only',  'TP ONLY': 'TP Only', 'THIRD PARTY': 'TP Only',
# #     'SAOD': 'SAOD',   'OD': 'SAOD',         'OWN DAMAGE': 'SAOD',
# #     'TW NEW': 'Comprehensive',   # new business TW → Comprehensive
# # }

# # # ── Subproduct  (motor only, name → id) ──────────────────────────────────────
# # df_sp_motor = df_subprod[df_subprod.get('product_name', pd.Series(dtype=str)).str.upper() == 'MOTOR'] \
# #               if 'product_name' in df_subprod.columns else df_subprod
# # subprod_id  = {str(r['sub_product_name']).strip(): si(r['sub_product_id'])
# #                for _, r in df_sp_motor.iterrows()}

# # # ── Vehicle type  (name → id, sub_product → default id) ──────────────────────
# # vt_name_id = {str(r['vehicle_type']).strip(): si(r['id']) for _, r in df_vtype.iterrows()}
# # vt_sub_default = {}   # sub_product_name → first vehicle_type_id for that sub
# # for _, r in df_vtype.iterrows():
# #     sp = str(r['sub_product_name']).strip()
# #     if sp not in vt_sub_default:
# #         vt_sub_default[sp] = si(r['id'])

# # # ── Fuel  (name → id) ─────────────────────────────────────────────────────────
# # fuel_id = {str(r['fuel_type']).strip().upper(): si(r['id']) for _, r in df_fuel_xl.iterrows()}
# # # Normalise varied fuel strings from MMV sheets → canonical names
# # _FUEL_NORM = {
# #     'P':'PETROL','PETROL':'PETROL',
# #     'D':'DIESEL','DIESEL':'DIESEL',
# #     'ELECTRIC':'ELECTRIC','E':'ELECTRIC','B':'ELECTRIC',
# #     'CNG':'CNG-LPG','LPG':'CNG-LPG','C':'CNG-LPG','CNG-LPG':'CNG-LPG',
# #     'CNG/PETROL':'CNG-LPG','HYBRID':'PETROL','PETROL/ELECTRIC':'ELECTRIC',
# # }
# # def fuel_lookup(raw):
# #     k = _FUEL_NORM.get(str(raw).strip().upper(), '')
# #     return fuel_id.get(k, -1), k

# # # =============================================================================
# # #  MMV INDEX BUILD
# # # =============================================================================
# # log("INFO", "Building MMV indexes …")
# # t1 = time.time()

# # # ── TW / Private Car MMV  (global — no company scope) ─────────────────────────
# # tw_make  = {}   # make_norm → (make_id, make_name)
# # tw_model = {}   # (make_norm, model_norm) → (model_id, model_name)
# # tw_var   = {}   # (make_norm, model_norm, var_norm) → (var_id, var_name, fuel_name, seating, cc, geared)
# # tw_makes_sorted = []          # sorted list for substring scanning
# # tw_models_of_make = {}        # make_norm → set of (model_norm, model_name, model_id)

# # for _, m in df_tw_mmv.iterrows():
# #     mkn = nt(m.get('vehicle_make',''));  mdn = nt(m.get('vehicle_model',''))
# #     vrn = nt(m.get('VariantDisplayName',''))
# #     mid = si(m.get('MakeID',-1)); mdid= si(m.get('ModelID',-1)); vid= si(m.get('VariantID',-1))
# #     fid_v, fn = fuel_lookup(m.get('Fuel',''))
# #     seat= si(m.get('SeatingCapacity',-1)); cc = si(m.get('CC',-1))
# #     gear= 1 if str(m.get('Is_Geared_Vehicle','')).upper() in ('TRUE','1','YES') else 0

# #     if mkn and mid != -1:
# #         tw_make.setdefault(mkn, (mid, str(m.get('vehicle_make','')).strip()))
# #     if mkn and mdn and mdid != -1:
# #         tw_model.setdefault((mkn, mdn), (mdid, str(m.get('vehicle_model','')).strip()))
# #         tw_models_of_make.setdefault(mkn, set()).add((mdn, str(m.get('vehicle_model','')).strip(), mdid))
# #     if mkn and mdn and vrn and vid != -1:
# #         tw_var.setdefault((mkn, mdn, vrn),
# #                           (vid, str(m.get('VariantDisplayName','')).strip(), fn, seat, cc, gear))

# # tw_makes_sorted = sorted([(k, tw_make[k][1], tw_make[k][0]) for k in tw_make],
# #                          key=lambda x: len(x[0]), reverse=True)

# # # ── CV MMV  (scoped by company_code) ──────────────────────────────────────────
# # cv_make  = {}   # (code, make_norm) → (make_id, make_name)
# # cv_model = {}   # (code, make_norm, model_norm) → (model_id, model_name)
# # cv_var   = {}   # (code, make_norm, model_norm, var_norm) → (var_id, var_name, fuel_name, seating)
# # cv_makes_of_co  = {}   # code → sorted list of (make_norm, make_name, make_id)
# # cv_models_of_mk = {}   # (code, make_norm) → set of (model_norm, model_name, model_id)
# # cv_models_of_co = {}   # code → {model_norm: [(model_id, model_name, make_id, make_name)]}

# # for _, m in df_cv_mmv.iterrows():
# #     code= nt(m.get('company_code',''))
# #     mkn = nt(m.get('vehicle_make',''));  mdn = nt(m.get('vehicle_model',''))
# #     vrn = nt(m.get('VariantName',''))
# #     mid = si(m.get('MakeID',-1)); mdid= si(m.get('ModelID',-1)); vid= si(m.get('VariantID',-1))
# #     fid_v, fn = fuel_lookup(m.get('Fuel',''))
# #     seat= si(m.get('SeatingCapacity',-1))

# #     if code and mkn and mid != -1:
# #         cv_make.setdefault((code, mkn), (mid, str(m.get('vehicle_make','')).strip()))
# #         cv_makes_of_co.setdefault(code, set()).add((mkn, str(m.get('vehicle_make','')).strip(), mid))
# #     if code and mkn and mdn and mdid != -1:
# #         cv_model.setdefault((code, mkn, mdn), (mdid, str(m.get('vehicle_model','')).strip()))
# #         cv_models_of_mk.setdefault((code, mkn), set()).add((mdn, str(m.get('vehicle_model','')).strip(), mdid))
# #         cv_models_of_co.setdefault(code, {}).setdefault(mdn, []).append(
# #             (mdid, str(m.get('vehicle_model','')).strip(), mid, str(m.get('vehicle_make','')).strip()))
# #     if code and mkn and mdn and vrn and vid != -1:
# #         cv_var.setdefault((code, mkn, mdn, vrn),
# #                           (vid, str(m.get('VariantName','')).strip(), fn, seat))

# # for k in list(cv_makes_of_co.keys()):
# #     cv_makes_of_co[k] = sorted(list(cv_makes_of_co[k]), key=lambda x: len(x[0]), reverse=True)

# # log("OK", (f"MMV indexes built in {time.time()-t1:.1f}s  |  "
# #            f"TW/PC makes={len(tw_make)} models={len(tw_model)} variants={len(tw_var)}  |  "
# #            f"CV make-rows={len(cv_make)} model-rows={len(cv_model)} variant-rows={len(cv_var)}"))

# # # =============================================================================
# # #  COMPANY SELECTION
# # # =============================================================================
# # print("\n" + "="*75)
# # print("AVAILABLE COMPANIES")
# # print("="*75)
# # for cid, row in sorted(company_by_id.items()):
# #     print(f"  {cid:3d}  |  {str(row['company_code']):20s}  |  {row['company_name']}")
# # print("="*75)

# # while True:
# #     try:
# #         CID      = int(input("\nEnter company_id: ").strip())
# #         CROW     = company_by_id[CID]
# #         CCODE    = str(CROW['company_code']).strip()
# #         CCODE_NT = nt(CCODE)
# #         log("OK", f"Company → {CROW['company_name']}  (id={CID}  code={CCODE})")
# #         break
# #     except (ValueError, KeyError):
# #         log("ERR", "Invalid company_id, try again.")

# # OUT_FILE = os.path.join(OUT_DIR, f"{CCODE}-Payin-Config.xlsx")

# # _api_key = os.getenv("OPENAI_API_KEY","").strip()
# # _model   = os.getenv("OPENAI_MODEL","gpt-4.1-mini")
# # if _api_key: log("OK",   f"OpenAI key found — model: {_model}")
# # else:        log("WARN", "No OPENAI_API_KEY — heuristic-only remark parsing")

# # # =============================================================================
# # #  PURE-PARSING HELPERS
# # # =============================================================================

# # def norm_pt(v):
# #     """Policy type raw string → COMP / TP / SAOD."""
# #     s = str(v).strip().upper()
# #     if s in ('COMP','COMPREHENSIVE'):        return 'COMP'
# #     if s in ('TP','TP ONLY','THIRD PARTY'):  return 'TP'
# #     if s in ('SAOD','OD','OWN DAMAGE'):      return 'SAOD'
# #     if s == 'TW NEW':                        return 'COMP'
# #     return 'COMP'

# # def vehicle_info(seg_text):
# #     """SEGMENT column → (sub_product_name, sub_product_id, default_vehicle_type_id)."""
# #     s = str(seg_text).strip().upper()
# #     # Two Wheeler
# #     if re.search(r'\bTW\b', s) or re.search(r'\b2W\b', s):
# #         sp = 'Two Wheeler'
# #         return sp, subprod_id.get(sp, -1), vt_name_id.get('TW Bike', vt_sub_default.get(sp, -1))
# #     # Private Car
# #     if re.search(r'\bPVT\s*CAR\b', s) or re.search(r'\bPRIVATE\s*CAR\b', s) or re.search(r'\bPC\b', s):
# #         sp = 'Private Car'
# #         return sp, subprod_id.get(sp, -1), vt_name_id.get('Private Car', vt_sub_default.get(sp, -1))
# #     # Passenger Vehicle
# #     if re.search(r'\bPCV\b', s) or re.search(r'\bPASSENGER\b', s):
# #         sp = 'Passenger Vehicle'
# #         return sp, subprod_id.get(sp, -1), vt_name_id.get('Auto rikshaw', vt_sub_default.get(sp, -1))
# #     # Miscellaneous — Tractor / Harvester
# #     if 'TRACTOR' in s or 'HARVESTER' in s:
# #         sp = 'Miscellaneous Vehicle'
# #         return sp, subprod_id.get(sp, -1), vt_name_id.get('Agriculture Tractor', vt_sub_default.get(sp, -1))
# #     # Miscellaneous — MISD / MISC
# #     if 'MISD' in s or 'MISC' in s:
# #         sp = 'Miscellaneous Vehicle'
# #         return sp, subprod_id.get(sp, -1), vt_name_id.get('Non Tractor', vt_sub_default.get(sp, -1))
# #     # Goods Vehicle — GCV / GVW / CV
# #     sp = 'Goods Vehicle'
# #     return sp, subprod_id.get(sp, -1), vt_name_id.get('Truck', vt_sub_default.get(sp, -1))

# # def parse_age(s):
# #     s = str(s).strip()
# #     if s.lower() in ('', 'nan', 'none', 'new', 'n/a'): return 0, 700
# #     su = s.upper()
# #     m = re.search(r'UPTO\s+(\d+)\s*(YEAR|YR|MONTH|MTH|MO)', su)
# #     if m: n, u = int(m.group(1)), m.group(2); return 0, n*12 if u.startswith('Y') else n
# #     m = re.search(r'(\d+)\s*TO\s*(\d+)\s*(YEAR|YR)?', su)
# #     if m: return int(m.group(1))*12, int(m.group(2))*12
# #     m = re.match(r'>\s*(\d+)\s*[-]\s*(\d+)(\+?)\s*[Yy]', s)
# #     if m: return int(m.group(1))*12+1, (700 if m.group(3)=='+' else int(m.group(2))*12)
# #     m = re.match(r'>\s*(\d+)\+?\s*[Yy]', s)
# #     if m: return int(m.group(1))*12+1, 700
# #     m = re.match(r'^(\d+)$', s.strip())
# #     if m: return 0, int(m.group(1))*12
# #     return 0, 700

# # def parse_cc(s):
# #     if not s or str(s).strip().lower() in ('', 'nan', 'none'): return 0, 99999, -1
# #     s = str(s).strip().upper().replace('CC','').strip()
# #     for pat, fn in [
# #         (r'^<\s*(\d+)$',                 lambda m: (0, int(m.group(1))-1, 1)),
# #         (r'^>\s*(\d+)$',                 lambda m: (int(m.group(1))+1, 99999, 1)),
# #         (r'^>\s*(\d+)\s*[-]\s*(\d+)$',  lambda m: (int(m.group(1))+1, int(m.group(2)), 1)),
# #         (r'^>=\s*(\d+)\s*[-]\s*(\d+)$', lambda m: (int(m.group(1)), int(m.group(2)), 1)),
# #         (r'^(\d+)\s*[-]\s*(\d+)$',      lambda m: (int(m.group(1)), int(m.group(2)), 1)),
# #         (r'^(\d+)$',                     lambda m: (int(m.group(1)), int(m.group(1)), 1)),
# #     ]:
# #         mt = re.match(pat, s)
# #         if mt: return fn(mt)
# #     return 0, 99999, -1

# # def parse_idv(text):
# #     su = str(text).upper()
# #     m = re.search(r'IDV\s+(?:UPTO|UP\s*TO|<=?)\s*([\d.]+)\s*(?:LAC|LAKH|L\b)', su)
# #     if m: return 1, 0.0, float(m.group(1))
# #     m = re.search(r'IDV\s+([\d.]+)\s*[-]\s*([\d.]+)\s*(?:LAC|LAKH)', su)
# #     if m: return 1, float(m.group(1)), float(m.group(2))
# #     return -1, 0.0, 0.0

# # def parse_weight(text):
# #     su = str(text).upper()
# #     m = re.search(r'([\d.]+)\s*KG\b', su)
# #     if m: return 1, 0, int(float(m.group(1)))
# #     m = re.search(r'UPTO\s+([\d.]+)\s*T(?:ON|ONNE)?', su)
# #     if m: return 1, 0, int(float(m.group(1))*1000)
# #     m = re.search(r'>\s*([\d.]+)\s*[-]\s*([\d.]+)\s*T(?:ON|ONNE)?', su)
# #     if m: return 1, int(float(m.group(1))*1000)+1, int(float(m.group(2))*1000)
# #     m = re.search(r'\b([\d.]+)\s*T(?:ON|ONNE)?\b', su)
# #     if m: return 1, 0, int(float(m.group(1))*1000)
# #     m = re.search(r'GVW\s+([\d.]+)', su)
# #     if m: v = float(m.group(1)); return 1, 0, int(v*1000) if v < 200 else int(v)
# #     return -1, 0, 99999

# # def parse_seating(text):
# #     su = str(text).upper()
# #     m = re.search(r'(?:SC|SEATING\s*CAPACITY|SEATING)\s*UPTO\s*(\d+)', su)
# #     if m: return 1, 1, int(m.group(1))
# #     m = re.search(r'(?:SC|SEATING\s*CAPACITY|SEATING)\s*(\d+)\s*[-]\s*(\d+)', su)
# #     if m: return 1, int(m.group(1)), int(m.group(2))
# #     m = re.search(r'(?:SC|SEATING\s*CAPACITY|SEATING)\s*(\d+)', su)
# #     if m: v = int(m.group(1)); return 1, v, v
# #     m = re.search(r'(\d+)\s*SEATER', su)
# #     if m: v = int(m.group(1)); return 1, v, v
# #     return -1, -1, -1

# # def parse_fuel_text(text):
# #     su = str(text).upper()
# #     if 'ELECTRIC' in su or ' EV ' in su: return fuel_id.get('ELECTRIC',-1), 'ELECTRIC'
# #     if 'CNG' in su or 'LPG' in su:       return fuel_id.get('CNG-LPG',-1),  'CNG-LPG'
# #     if 'DIESEL' in su:                   return fuel_id.get('DIESEL',-1),    'DIESEL'
# #     if 'PETROL' in su:                   return fuel_id.get('PETROL',-1),    'PETROL'
# #     return -1, ''

# # def ncb_flag(text):
# #     su = str(text).upper()
# #     if any(x in su for x in ('WITHOUT NCB','W/O NCB','NON NCB','NON-NCB','ZERO NCB')): return 0
# #     if 'NCB' in su: return 1
# #     return -1

# # def irda_flag(text):
# #     su = str(text).upper()
# #     return 1 if any(x in su for x in ('IRDA TP','IRDA RATE','IRDA')) else -1

# # def cpa_flag(text):
# #     su = str(text).upper()
# #     if 'CPA' in su and any(x in su for x in ('INCLUD','WITH CPA')): return 1
# #     return -1

# # def keep_included(text):
# #     """Return only the 'included' portion — strip excluded/except/rejected chunks."""
# #     s = str(text).strip()
# #     if not s or s.lower() in ('nan', 'none'): return ''
# #     kept = []
# #     for chunk in re.split(r';', s):
# #         up = chunk.upper()
# #         if any(t in up for t in ('DECLIN','REJECT')): continue
# #         for tok in (' BUT ',' EXCEPT ',' EXCLUDE ',' OTHER THAN ',' NOT CONSIDER'):
# #             idx = up.find(tok)
# #             if idx != -1: chunk = chunk[:idx]; break
# #         chunk = chunk.strip()
# #         if chunk: kept.append(chunk)
# #     return ' '.join(x for x in kept if x)

# # def is_pure_exclusion(text):
# #     """True if the remark has ONLY exclusion language and no inclusion items."""
# #     s = str(text).strip().upper()
# #     if not s or s in ('NAN','NONE'): return False
# #     EXCL = ('EXCEPT','EXCLUDE','OTHER THAN','REJECT','DECLIN','NOT CONSIDER',
# #             'HR 68','EXCLUDED')
# #     INCL = ('ONLY','INCLUD','HONDA','BAJAJ','TATA','MARUTI','HYUNDAI','KIA',
# #             'MAHINDRA','IDV','NCB','ZONE','BRANCH')
# #     if any(t in s for t in INCL): return False
# #     return any(t in s for t in EXCL)

# # # =============================================================================
# # #  MMV RESOLUTION
# # # =============================================================================

# # def resolve_tw(raw_make, raw_model, raw_variant, remark):
# #     mkn = nt(raw_make) if raw_make else ''
# #     mdn = nt(raw_model) if raw_model else ''
# #     vrn = nt(raw_variant) if raw_variant else ''
# #     mk_id=-1; mk_name=''; md_id=-1; md_name=''; vr_id=-1; vr_name=''
# #     fuel=''; seat=-1; cc=-1; gear=-1
# #     inc = nt(keep_included(remark))

# #     # 1. Make
# #     if mkn:
# #         if mkn in tw_make:
# #             mk_id, mk_name = tw_make[mkn]
# #         else:
# #             mw = set(mkn.split())
# #             for cn, cname, cid in tw_makes_sorted:
# #                 if mkn in cn or cn in mkn or mw.issubset(set(cn.split())):
# #                     mk_id, mk_name, mkn = cid, cname, cn; break
# #         if mk_id == -1:  # scan remark
# #             for cn, cname, cid in tw_makes_sorted:
# #                 if cn and re.search(rf'\b{re.escape(cn)}\b', inc):
# #                     mk_id, mk_name, mkn = cid, cname, cn; break

# #     # 2. Model (only if mentioned)
# #     if mdn and mk_id != -1:
# #         if (mkn, mdn) in tw_model:
# #             md_id, md_name = tw_model[(mkn, mdn)]
# #         else:
# #             for cmn, cmname, cmid in sorted(tw_models_of_make.get(mkn, set()),
# #                                             key=lambda x: len(x[0]), reverse=True):
# #                 if mdn in cmn or cmn in mdn:
# #                     md_id, md_name, mdn = cmid, cmname, cmn; break

# #     # 3. Variant (only if mentioned)
# #     if vrn and mk_id != -1 and md_id != -1:
# #         key = (mkn, mdn, vrn)
# #         if key in tw_var:
# #             vr_id, vr_name, fuel, seat, cc, gear = tw_var[key]
# #         else:
# #             for (cmn,cmdn,cvn), (vid,vname,vf,vs,vc,vg) in tw_var.items():
# #                 if cmn==mkn and cmdn==mdn and (vrn in cvn or cvn in vrn):
# #                     vr_id,vr_name,fuel,seat,cc,gear = vid,vname,vf,vs,vc,vg; break

# #     # Infer fuel/seating from first matching variant when no variant given
# #     if mk_id != -1 and md_id != -1 and not vrn:
# #         for (cmn,cmdn,_),(vid,vname,vf,vs,vc,vg) in tw_var.items():
# #             if cmn==mkn and cmdn==mdn and vf:
# #                 fuel,seat,cc,gear = vf,vs,vc,vg; break

# #     return mk_id, mk_name, md_id, md_name, vr_id, vr_name, fuel, seat, cc, gear


# # def resolve_cv(raw_make, raw_model, raw_variant, remark):
# #     code = CCODE_NT
# #     mkn  = nt(raw_make)    if raw_make    else ''
# #     mdn  = nt(raw_model)   if raw_model   else ''
# #     vrn  = nt(raw_variant) if raw_variant else ''
# #     mk_id=-1; mk_name=''; md_id=-1; md_name=''; vr_id=-1; vr_name=''
# #     fuel=''; seat=-1
# #     inc = nt(keep_included(remark))

# #     # 1. Make
# #     if mkn:
# #         if (code, mkn) in cv_make:
# #             mk_id, mk_name = cv_make[(code, mkn)]
# #         else:
# #             mw = set(mkn.split())
# #             for cn, cname, cid in cv_makes_of_co.get(code, []):
# #                 if mkn in cn or cn in mkn or mw.issubset(set(cn.split())):
# #                     mk_id, mk_name, mkn = cid, cname, cn; break
# #         if mk_id == -1:  # scan remark
# #             for cn, cname, cid in cv_makes_of_co.get(code, []):
# #                 if cn and re.search(rf'\b{re.escape(cn)}\b', inc):
# #                     mk_id, mk_name, mkn = cid, cname, cn; break

# #     # 2. Model
# #     if mdn:
# #         if mk_id != -1 and (code, mkn, mdn) in cv_model:
# #             md_id, md_name = cv_model[(code, mkn, mdn)]
# #         elif mk_id != -1:
# #             for cmn,cmname,cmid in sorted(cv_models_of_mk.get((code, mkn), set()),
# #                                           key=lambda x: len(x[0]), reverse=True):
# #                 if mdn in cmn or cmn in mdn:
# #                     md_id, md_name, mdn = cmid, cmname, cmn; break
# #         if md_id == -1:
# #             hits = cv_models_of_co.get(code, {}).get(mdn, [])
# #             if hits: md_id, md_name, mk_id, mk_name = hits[0]

# #     # 3. Variant
# #     if vrn and mk_id != -1 and md_id != -1:
# #         key = (code, mkn, mdn, vrn)
# #         if key in cv_var:
# #             vr_id, vr_name, fuel, seat = cv_var[key]
# #         else:
# #             for (ck,cmn,cmdn,cvn),(vid,vname,vf,vs) in cv_var.items():
# #                 if ck==code and cmn==mkn and cmdn==mdn and (vrn in cvn or cvn in vrn):
# #                     vr_id,vr_name,fuel,seat = vid,vname,vf,vs; break

# #     if mk_id != -1 and md_id != -1 and not vrn:
# #         for (ck,cmn,cmdn,_),(vid,vname,vf,vs) in cv_var.items():
# #             if ck==code and cmn==mkn and cmdn==mdn and vf:
# #                 fuel, seat = vf, vs; break

# #     return mk_id, mk_name, md_id, md_name, vr_id, vr_name, fuel, seat

# # # =============================================================================
# # #  OPENAI REMARK PARSER
# # # =============================================================================

# # def heuristic(remark, seg, pt):
# #     su = str(remark).upper()
# #     fid,fn  = parse_fuel_text(remark)
# #     iw,fw,tw_= parse_weight(remark)
# #     iss,fs,ts= parse_seating(remark)
# #     ii,fi,ti = parse_idv(remark)
# #     return {
# #         'vehicle_make':'','vehicle_model':'','vehicle_variant':'',
# #         'is_with_ncb':ncb_flag(su), 'is_irda_tp':irda_flag(su),
# #         'is_cpa_included':cpa_flag(su),
# #         'fuel_type':fn,
# #         'seating_cap':-1,'from_seating':fs,'to_seating':ts,
# #         'is_weight':iw,'from_weight_kg':fw,'to_weight_kg':tw_,
# #         'idv_cap':ii,'from_idv':fi,'to_idv':ti,
# #         'is_cc':-1,'from_cc':0,'to_cc':99999,
# #     }

# # _nc=0; _nh=0; _ne=0; _ms=0

# # def parse_remark(remark, co_name, seg, pt, row_n=0):
# #     global _nc, _ne, _ms
# #     ak = os.getenv("OPENAI_API_KEY","").strip()
# #     if not ak: return heuristic(remark, seg, pt)

# #     included = keep_included(remark)
# #     short    = (remark[:80]+'…') if len(remark)>80 else remark
# #     log("API", f"Row {row_n:>4} | call #{_nc+1} | seg={seg!r:20} | {short!r}")

# #     prompt = f"""You are an expert Indian motor insurance data extractor.
# # Analyse the remark carefully. Return ONLY a valid JSON object — no markdown, no preamble — with EXACTLY these keys:

# #   vehicle_make    : INCLUDED makes comma-separated (e.g. "HONDA,HYUNDAI,KIA"). Empty if none.
# #   vehicle_model   : INCLUDED models comma-separated. Empty if none.
# #   vehicle_variant : Specific variant name if mentioned, else "".
# #   is_with_ncb     : 1 if NCB cases INCLUDED. 0 if WITHOUT/NON/ZERO NCB. -1 if not mentioned.
# #   is_irda_tp      : 1 if IRDA TP rate mentioned, -1 otherwise.
# #   is_cpa_included : 1 if CPA included/mentioned, -1 otherwise.
# #   fuel_type       : DIESEL | PETROL | ELECTRIC | CNG-LPG | "" (empty if not mentioned).
# #   seating_cap     : exact integer if single seating value given, -1 otherwise.
# #   from_seating    : lower seating bound (int), -1 if N/A.
# #   to_seating      : upper seating bound (int), -1 if N/A.
# #   is_weight       : 1 if GVW/weight/tonnage mentioned, -1 otherwise.
# #   from_weight_kg  : lower weight KG int, 0 if N/A.
# #   to_weight_kg    : upper weight KG int, 99999 if N/A.
# #   idv_cap         : 1 if IDV cap mentioned, -1 otherwise.
# #   from_idv        : lower IDV in Lacs (float), 0 if N/A.
# #   to_idv          : upper IDV in Lacs (float), 0 if N/A.
# #   is_cc           : 1 if engine CC/capacity mentioned, -1 otherwise.
# #   from_cc         : lower CC int, 0 if N/A.
# #   to_cc           : upper CC int, 99999 if N/A.

# # RULES — READ CAREFULLY:
# # 1. ONLY list makes/models that are INCLUDED. IGNORE anything after EXCEPT / BUT / EXCLUDE /
# #    OTHER THAN / REJECT / DECLINE / ONLY EXCEPT.
# # 2. If remark has ONLY exclusion text (e.g. "Except TATA" or "HR 68 EXCLUDED"), set
# #    vehicle_make="" and vehicle_model="".
# # 3. SC = Seating Capacity. "SC 7" → seating_cap=7. "SC upto 7" → from_seating=1, to_seating=7.
# # 4. Tons → KG: 1 ton = 1000 KG. "7.5T" → to_weight_kg=7500.
# # 5. "IDV upto 10 lacs" → idv_cap=1, from_idv=0, to_idv=10.
# # 6. "Upto 1500 CC" → is_cc=1, from_cc=0, to_cc=1500.
# # 7. TRACTOR segment → vehicle_variant="Agriculture Tractor".

# # company: {co_name}
# # segment: {seg}
# # policy_type: {pt}
# # remark_original: {remark}
# # remark_included_only: {included}
# # """
# #     body = {"model": _model,
# #             "messages": [
# #                 {"role":"system","content":"Return only valid JSON, no markdown, no explanation."},
# #                 {"role":"user","content":prompt}],
# #             "response_format":{"type":"json_object"}, "temperature":0}
# #     req = urllib.request.Request(
# #         "https://api.openai.com/v1/chat/completions",
# #         data=json.dumps(body).encode(),
# #         headers={"Content-Type":"application/json","Authorization":f"Bearer {ak}"},
# #         method="POST")
# #     ts = time.time()
# #     try:
# #         with urllib.request.urlopen(req, timeout=30) as r:
# #             p = json.loads(json.loads(r.read())["choices"][0]["message"]["content"])
# #         ms = int((time.time()-ts)*1000); _nc+=1; _ms+=ms
# #         def _s(k,d=""): return str(p.get(k,d)).strip()
# #         def _i(k,d=-1):
# #             try: return int(p.get(k,d))
# #             except: return d
# #         def _f(k,d=0.0):
# #             try: return float(p.get(k,d))
# #             except: return d
# #         result = {
# #             'vehicle_make':_s('vehicle_make'),'vehicle_model':_s('vehicle_model'),
# #             'vehicle_variant':_s('vehicle_variant'),
# #             'is_with_ncb':_i('is_with_ncb'),'is_irda_tp':_i('is_irda_tp'),
# #             'is_cpa_included':_i('is_cpa_included'),
# #             'fuel_type':_s('fuel_type').upper(),
# #             'seating_cap':_i('seating_cap'),'from_seating':_i('from_seating'),
# #             'to_seating':_i('to_seating'),
# #             'is_weight':_i('is_weight'),
# #             'from_weight_kg':_i('from_weight_kg',0),'to_weight_kg':_i('to_weight_kg',99999),
# #             'idv_cap':_i('idv_cap'),'from_idv':_f('from_idv'),'to_idv':_f('to_idv'),
# #             'is_cc':_i('is_cc'),'from_cc':_i('from_cc',0),'to_cc':_i('to_cc',99999),
# #         }
# #         log("OK", (f"Row {row_n:>4} | {ms}ms | "
# #                    f"make={result['vehicle_make']!r:15} model={result['vehicle_model']!r:12} "
# #                    f"ncb={result['is_with_ncb']} irda={result['is_irda_tp']} "
# #                    f"cc={result['is_cc']} fuel={result['fuel_type']!r}"))
# #         return result
# #     except urllib.error.HTTPError as e:
# #         _ne+=1; log("ERR", f"Row {row_n:>4} | HTTP {e.code} → heuristic")
# #     except Exception as e:
# #         _ne+=1; log("ERR", f"Row {row_n:>4} | {e} → heuristic")
# #     return heuristic(remark, seg, pt)

# # # =============================================================================
# # #  OUTPUT COLUMN ORDER
# # # =============================================================================
# # COLS = [
# #     'id','company_id','company_code','segment_id','segment',
# #     'subproduct_id','sub_product_name','lob_id','lob_name',
# #     'business_type_id','business_type','is_highend_lob',
# #     'rto_group_id','rto_group_name',
# #     'payin_od_rate','payin_tp_rate','payout_od_rate','payout_tp_rate',
# #     'extra_tp_rate','eff_from_date','eff_to_date',
# #     'fuel_type_id','fuel_type',
# #     'is_on_net','is_one_year_pay_on_newbusiness','is_cpa_included',
# #     'is_geared_vehicle','is_cc_considered','from_cc','to_cc',
# #     'is_premium_considered','from_premium','to_premium',
# #     'is_mmv_considered','make_id','vehicle_make','model_id','vehicle_model',
# #     'variant_id','vehicle_variant',
# #     'is_seating_cap_consider','from_seating_cap','to_seating_cap',
# #     'is_no_of_wheel_consider','from_no_of_wheel','to_no_of_wheel',
# #     'vehicle_type_id','ppi_in','ppi_out',
# #     'is_irda_tp_included','is_longterm_renewal_pay',
# #     'is_weightage_considered','from_weightage_kg','to_weightage_kg',
# #     'is_nil_dep_considered','is_organization_type',
# #     'from_age_month','to_age_month',
# #     'is_with_ncb','is_idv_cap_consider','from_idv','to_idv',
# #     'is_breakin_consider','is_active',
# # ]

# # def make_row(
# #     sid,sname,spid,spname,rto_id,rto_name,
# #     pod,ptp,pod2,ptp2,
# #     fid,fname,is_on_net,cpa,geared,
# #     iscc,fcc,tcc,
# #     iswt,fwt,twt,
# #     vtid,
# #     is_mmv,mkid,mkname,mdid,mdname,vrid,vrname,
# #     issc,fsc,tsc,
# #     nw,fnw,tnw,
# #     fage,tage,
# #     ncb,irda,
# #     idvcap,fidv,tidv,
# # ):
# #     return {
# #         'id':0, 'company_id':CID, 'company_code':CCODE,
# #         'segment_id':sid, 'segment':sname,
# #         'subproduct_id':spid, 'sub_product_name':spname,
# #         'lob_id':-1, 'lob_name':'',
# #         'business_type_id':-1, 'business_type':'Not Considered', 'is_highend_lob':False,
# #         'rto_group_id':rto_id, 'rto_group_name':rto_name,
# #         'payin_od_rate':pod, 'payin_tp_rate':ptp,
# #         'payout_od_rate':pod2, 'payout_tp_rate':ptp2,
# #         'extra_tp_rate':0, 'eff_from_date':'2026-01-01', 'eff_to_date':'2026-01-16',
# #         'fuel_type_id':fid, 'fuel_type':fname,
# #         'is_on_net':is_on_net,
# #         'is_one_year_pay_on_newbusiness':-1, 'is_cpa_included':cpa,
# #         'is_geared_vehicle':geared,
# #         'is_cc_considered':iscc, 'from_cc':fcc, 'to_cc':tcc,
# #         'is_premium_considered':-1, 'from_premium':-1, 'to_premium':-1,
# #         'is_mmv_considered':is_mmv,
# #         'make_id':mkid, 'vehicle_make':mkname,
# #         'model_id':mdid, 'vehicle_model':mdname,
# #         'variant_id':vrid, 'vehicle_variant':vrname,
# #         'is_seating_cap_consider':issc, 'from_seating_cap':fsc, 'to_seating_cap':tsc,
# #         'is_no_of_wheel_consider':nw, 'from_no_of_wheel':fnw, 'to_no_of_wheel':tnw,
# #         'vehicle_type_id':vtid, 'ppi_in':0, 'ppi_out':0,
# #         'is_irda_tp_included':irda, 'is_longterm_renewal_pay':-1,
# #         'is_weightage_considered':iswt, 'from_weightage_kg':fwt, 'to_weightage_kg':twt,
# #         'is_nil_dep_considered':-1, 'is_organization_type':-1,
# #         'from_age_month':fage, 'to_age_month':tage,
# #         'is_with_ncb':ncb,
# #         'is_idv_cap_consider':idvcap, 'from_idv':fidv, 'to_idv':tidv,
# #         'is_breakin_consider':-1, 'is_active':True,
# #     }

# # # =============================================================================
# # #  PROCESS ONE INPUT FILE
# # # =============================================================================

# # def process(input_file):
# #     global _nh
# #     log("INFO", f"Reading: {input_file}")
# #     tr = time.time()
# #     df = pd.read_excel(input_file)
# #     df.columns = [c.strip() for c in df.columns]
# #     log("OK", f"Loaded {len(df)} rows in {time.time()-tr:.1f}s | cols: {list(df.columns)}")

# #     def col(*names):
# #         for n in names:
# #             if n in df.columns: return n
# #         return None

# #     c_seg  = col('SEGMENT','Segment')
# #     c_pt   = col('POLICY TYPE','Policy Type','POLICYTYPE')
# #     c_loc  = col('LOCATION','Location')
# #     c_pay  = col('PAYIN','Payin')
# #     c_pout = col('PAYOUT','Payout','Calculated Payout')
# #     c_rem  = col('REMARK','Remark','REMARKS','CALCULATION EXPLANATION')
# #     c_age  = col('AGE','Age','AGE BAND')
# #     c_cc   = col('CC BAND','CC Band','CC')
# #     c_tw   = col('TW TYPE','TW Type')
# #     c_co   = col('COMPANY NAME','Company Name','COMPANY')

# #     log("INFO", f"col-map → seg={c_seg!r} pt={c_pt!r} loc={c_loc!r} "
# #         f"pay={c_pay!r} pout={c_pout!r} rem={c_rem!r} age={c_age!r}")

# #     # estimate API calls
# #     uniq_rem = set()
# #     if c_rem:
# #         for v in df[c_rem].fillna('').astype(str): uniq_rem.add(v.strip())
# #     log("INFO", f"Unique remarks: {len(uniq_rem)} → "
# #         + (f"~{len(uniq_rem)} API calls" if _api_key else "heuristic only"))

# #     out_rows = []; cache = {}; total = len(df); tp = time.time()
# #     print(f"\n  {'='*58}\n  Processing {total} rows …\n  {'='*58}\n")

# #     for idx, (_, row) in enumerate(df.iterrows(), 1):
# #         bar(idx, total)

# #         if idx == 1 or idx % 25 == 0 or idx == total:
# #             el = time.time()-tp; rate = idx/el if el else 0; eta = (total-idx)/rate if rate else 0
# #             avg = (_ms/_nc) if _nc else 0
# #             log("INFO", (f"Row {idx:>4}/{total} | {el:.0f}s | ETA {eta:.0f}s | "
# #                          f"rate={rate:.1f}/s | api={_nc} avg={avg:.0f}ms | "
# #                          f"cache={_nh} err={_ne} | out={len(out_rows)}"))

# #         def g(c, d=''):
# #             if c is None: return d
# #             v = row.get(c, d)
# #             return v if v is not None and str(v).strip() not in ('nan','None','NaN') else d

# #         # ── Policy type & segment ─────────────────────────────────────────────
# #         pt_raw = str(g(c_pt, 'COMP')).strip()
# #         pt     = norm_pt(pt_raw)
# #         seg_nm = PT_SEG.get(pt, 'Comprehensive')
# #         seg_id = seg_name_to_id.get(seg_nm, 1)

# #         # ── Rates ─────────────────────────────────────────────────────────────
# #         payin  = sf(g(c_pay, 0))
# #         payout = sf(g(c_pout, 0))
# #         if   pt == 'TP':   pod, ptp, pod2, ptp2 = 0, payin, 0, payout
# #         elif pt == 'SAOD': pod, ptp, pod2, ptp2 = payin, 0, payout, 0
# #         else:              pod = ptp = payin; pod2 = ptp2 = payout

# #         # ── Vehicle / subproduct ──────────────────────────────────────────────
# #         seg_text = str(g(c_seg, '')).strip()
# #         spname, spid, vtid = vehicle_info(seg_text)
# #         is_cv   = spname in ('Goods Vehicle','Passenger Vehicle','Miscellaneous Vehicle')

# #         # ── Location (rto_group_id always 0 per corrections.txt) ─────────────
# #         rto_name = str(g(c_loc, '')).strip()
# #         rto_id   = 0

# #         # ── Other columns ─────────────────────────────────────────────────────
# #         remark   = str(g(c_rem, '')).strip()
# #         co_name  = str(g(c_co, CCODE)).strip()
# #         fage, tage = parse_age(str(g(c_age, '')))
# #         fcc0, tcc0, iscc0 = parse_cc(str(g(c_cc, '')))

# #         # TW geared / vehicle-type override
# #         tw_raw = str(g(c_tw, '')).strip().lower()
# #         geared = -1
# #         if spname == 'Two Wheeler' and tw_raw:
# #             if 'scooter'  in tw_raw: vtid = vt_name_id.get('TW Scooter', vtid); geared = 0
# #             elif 'bike'   in tw_raw: vtid = vt_name_id.get('TW Bike',    vtid); geared = 1
# #             elif 'electric' in tw_raw: vtid = vt_name_id.get('TW Electric Bike', vtid); geared = -1

# #         # ── is_on_net ─────────────────────────────────────────────────────────
# #         is_on_net = True if pt == 'COMP' else False

# #         # ── OpenAI / cache ────────────────────────────────────────────────────
# #         ck = (remark, co_name, seg_text, pt)
# #         if ck in cache:
# #             _nh += 1
# #             log("CACHE", f"Row {idx:>4} | HIT #{_nh} | {remark[:50]!r}")
# #             meta = cache[ck]
# #         else:
# #             meta  = parse_remark(remark, co_name, seg_text, pt, idx)
# #             cache[ck] = meta

# #         # ── Pull meta fields ──────────────────────────────────────────────────
# #         ncb    = meta['is_with_ncb']
# #         irda   = meta['is_irda_tp']
# #         cpa    = meta['is_cpa_included']
# #         raw_mk = meta['vehicle_make']
# #         raw_md = meta['vehicle_model']
# #         raw_vr = meta['vehicle_variant']

# #         # Fuel
# #         m_fuel = meta.get('fuel_type', '')
# #         if m_fuel and m_fuel in fuel_id: fid_v, fn_v = fuel_id[m_fuel], m_fuel
# #         else: fid_v, fn_v = parse_fuel_text(remark)

# #         # Seating
# #         msc  = meta.get('seating_cap',-1)
# #         mfsc = meta.get('from_seating',-1)
# #         mtsc = meta.get('to_seating',-1)
# #         if msc != -1:                        issc_v,fsc_v,tsc_v = 1,msc,msc
# #         elif mfsc != -1 or mtsc != -1:       issc_v=1; fsc_v=mfsc if mfsc!=-1 else 1; tsc_v=mtsc if mtsc!=-1 else 99
# #         else:                                issc_v,fsc_v,tsc_v = parse_seating(remark)

# #         # Weight
# #         m_wt  = meta.get('is_weight',-1)
# #         if m_wt == 1: iswt_v,fwt_v,twt_v = 1,meta['from_weight_kg'],meta['to_weight_kg']
# #         else:         iswt_v,fwt_v,twt_v = parse_weight(remark)

# #         # IDV
# #         m_idv = meta.get('idv_cap',-1)
# #         if m_idv == 1: idv_v,fidv_v,tidv_v = 1,meta['from_idv'],meta['to_idv']
# #         else:          idv_v,fidv_v,tidv_v = parse_idv(remark)

# #         # CC  (column-level overrides OpenAI if column exists)
# #         iscc_v, fcc_v, tcc_v = iscc0, fcc0, tcc0
# #         if iscc_v == -1 and meta.get('is_cc',-1) == 1:
# #             iscc_v, fcc_v, tcc_v = 1, meta['from_cc'], meta['to_cc']

# #         # Pure exclusion → clear MMV
# #         if is_pure_exclusion(remark):
# #             raw_mk = raw_md = raw_vr = ''
# #             log("INFO", f"Row {idx:>4} | pure-exclusion remark → MMV cleared")

# #         # ── Expand multiple makes ─────────────────────────────────────────────
# #         make_list = [m.strip() for m in re.split(r'[,&]+', raw_mk) if m.strip()] or ['']
# #         if len(make_list) > 1:
# #             log("INFO", f"Row {idx:>4} | expanding {len(make_list)} makes: {make_list}")

# #         for one_make in make_list:
# #             if is_cv:
# #                 mkid,mkname,mdid,mdname,vrid,vrname,i_fuel,i_seat = \
# #                     resolve_cv(one_make, raw_md, raw_vr, remark)
# #                 i_cc=-1; i_gear=-1
# #             else:
# #                 mkid,mkname,mdid,mdname,vrid,vrname,i_fuel,i_seat,i_cc,i_gear = \
# #                     resolve_tw(one_make, raw_md, raw_vr, remark)

# #             if one_make and mkid == -1:
# #                 log("WARN", f"Row {idx:>4} | make '{one_make}' NOT in MMV for {CCODE}")

# #             # Tractor → vehicle_variant override
# #             if spname == 'Miscellaneous Vehicle' and 'TRACTOR' in seg_text.upper():
# #                 if not vrname: vrname = 'Agriculture Tractor'

# #             is_mmv = 1 if (mkid!=-1 or mdid!=-1 or vrid!=-1 or
# #                            one_make or raw_md or raw_vr) else -1

# #             # Fuel fallback from MMV variant data
# #             fin_fid, fin_fn = fid_v, fn_v
# #             if fin_fid == -1 and i_fuel and i_fuel in fuel_id:
# #                 fin_fid = fuel_id[i_fuel]; fin_fn = i_fuel

# #             # Seating fallback from MMV
# #             fin_issc, fin_fsc, fin_tsc = issc_v, fsc_v, tsc_v
# #             if fin_issc == -1 and i_seat > 0:
# #                 fin_issc=1; fin_fsc=i_seat; fin_tsc=i_seat

# #             # Geared fallback from MMV variant
# #             fin_gear = geared
# #             if spname == 'Two Wheeler' and fin_gear == -1 and i_gear != -1:
# #                 fin_gear = i_gear

# #             out_rows.append(make_row(
# #                 seg_id, seg_nm, spid, spname, rto_id, rto_name,
# #                 pod, ptp, pod2, ptp2,
# #                 fin_fid, fin_fn, is_on_net, cpa, fin_gear,
# #                 iscc_v, fcc_v, tcc_v,
# #                 iswt_v, fwt_v, twt_v,
# #                 vtid,
# #                 is_mmv,
# #                 mkid, mkname if mkname else one_make,
# #                 mdid, mdname if mdname else raw_md,
# #                 vrid, vrname if vrname else raw_vr,
# #                 fin_issc, fin_fsc, fin_tsc,
# #                 -1, -1, -1,
# #                 fage, tage,
# #                 ncb, irda,
# #                 idv_v, fidv_v, tidv_v,
# #             ))

# #     el = time.time()-tp; avg = (_ms/_nc) if _nc else 0
# #     print()
# #     log("OK","="*60)
# #     log("OK",f"DONE  input={total}  output={len(out_rows)}  time={el:.1f}s ({el/60:.1f}min)")
# #     log("OK",f"  API calls={_nc}  avg={avg:.0f}ms  cache={_nh}  errors={_ne}")
# #     log("OK","="*60)
# #     return pd.DataFrame(out_rows)

# # # =============================================================================
# # #  MAIN LOOP  — process one or more files, all appended to same output
# # # =============================================================================
# # input_file = input("\nEnter Shriram input Excel file path: ").strip().strip('"')

# # while True:
# #     try:
# #         df_out = process(input_file)
# #         df_out = df_out[[c for c in COLS if c in df_out.columns]]

# #         if os.path.exists(OUT_FILE):
# #             log("INFO", f"Appending to existing: {OUT_FILE}")
# #             df_out = pd.concat([pd.read_excel(OUT_FILE), df_out], ignore_index=True)

# #         df_out.to_excel(OUT_FILE, index=False)
# #         log("OK", f"Saved → {OUT_FILE}  ({len(df_out)} rows)")
# #         log("OK", f"Log   → {_log_path}")

# #     except Exception as e:
# #         import traceback
# #         log("ERR", f"FATAL: {e}")
# #         traceback.print_exc()

# #     print("\n" + "="*75)
# #     print("  1  Process another Shriram file (appends to same output)")
# #     print("  2  Exit")
# #     ch = input("Choice: ").strip()
# #     if ch == "2":
# #         log("OK","Goodbye!")
# #         if _log_file: _log_file.close()
# #         break
# #     input_file = input("Next Shriram file path: ").strip().strip('"')

# import pandas as pd
# import re, sys, os, json, time, urllib.request, urllib.error
# from datetime import datetime

# # =============================================================================
# #  SHRIRAM GENERAL INSURANCE — PayinConfig Generator  v3
# #  Reads master data from Masters_-_dec_2025.xlsx  +  JSON reference files
# #  Run:  python SHRIRAM_PayinConfig_v3.py
# # =============================================================================

# # ── Logging ──────────────────────────────────────────────────────────────────
# C = {
#     "OK"   : "\033[32m", "WARN" : "\033[33m", "ERR"  : "\033[31m",
#     "API"  : "\033[36m", "CACHE": "\033[90m", "INFO" : "",
# }
# _log_file = None

# def log(level, msg):
#     ts   = datetime.now().strftime("%H:%M:%S")
#     line = f"[{ts}][{level}] {msg}"
#     col  = C.get(level, "")
#     print(f"{col}{line}\033[0m" if col else line)
#     if _log_file:
#         _log_file.write(line + "\n")
#         _log_file.flush()

# def bar(cur, tot, w=48):
#     pct = cur / tot if tot else 0
#     b   = "\u2588" * int(w * pct) + "\u2591" * (w - int(w * pct))
#     print(f"\r  [{b}] {cur}/{tot} ({pct*100:.1f}%)", end="", flush=True)
#     if cur == tot: print()

# # ── Tiny helpers ──────────────────────────────────────────────────────────────
# def nt(v):
#     """Normalise: UPPER + collapse non-alnum to single space."""
#     return re.sub(r'\s+', ' ', re.sub(r'[^A-Z0-9]+', ' ', str(v).upper())).strip()

# def si(v, d=-1):
#     try:   return int(float(str(v)))
#     except: return d

# def sf(v, d=0.0):
#     if isinstance(v, str): v = v.strip().replace('%','')
#     try:   return float(v)
#     except: return d

# def load_dotenv(p):
#     if not os.path.exists(p): return
#     for line in open(p, encoding="utf-8"):
#         s = line.strip()
#         if not s or s.startswith('#') or '=' not in s: continue
#         k, v = s.split('=', 1)
#         k = k.strip(); v = v.strip().strip('"').strip("'")
#         if k and k not in os.environ: os.environ[k] = v

# # =============================================================================
# #  STARTUP
# # =============================================================================
# print("\n" + "="*75)
# print("  SHRIRAM GENERAL INSURANCE — PayinConfig Generator  v3")
# print("="*75)

# MASTERS = input("Path to Masters_-_dec_2025.xlsx  : ").strip().strip('"')
# JSON_DIR= input("Path to JSON files folder        : ").strip().strip('"')
# OUT_DIR = input("Output folder path               : ").strip().strip('"')
# os.makedirs(OUT_DIR, exist_ok=True)

# _log_path = os.path.join(OUT_DIR, f"shriram_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
# try:
#     _log_file = open(_log_path, "w", encoding="utf-8")
#     log("OK", f"Log → {_log_path}")
# except Exception as e:
#     print(f"[WARN] Cannot open log file: {e}")

# load_dotenv(os.path.join(JSON_DIR, ".env"))
# load_dotenv(".env")

# # =============================================================================
# #  LOAD MASTERS  (Excel sheets take priority; JSON files fill any gaps)
# # =============================================================================
# log("INFO", f"Loading master data from Excel: {MASTERS}")
# t0  = time.time()
# xl  = pd.ExcelFile(MASTERS)

# def rsh(name):
#     df = pd.read_excel(xl, sheet_name=name)
#     df.columns = [str(c).strip() for c in df.columns]
#     return df

# df_company = rsh("Insurance Company Master")   # company_id, company_code, company_name
# df_segment = rsh("Segment Master")            # id, segment
# df_subprod = rsh("SubProduct Master")         # sub_product_id, sub_product_name (+product_name)
# df_vtype   = rsh("Vehicle type")              # id, vehicle_type, no_of_wheels, sub_product_name
# df_fuel_xl = rsh("fuel master")               # id, fuel_type
# df_tw_mmv  = rsh("tw_pc_mmv master")          # MakeID/vehicle_make/ModelID/vehicle_model/VariantID/VariantDisplayName/CC/Fuel/SeatingCapacity/Is_Geared_Vehicle
# df_cv_mmv  = rsh("cv_mmv master")             # company_code/MakeID/vehicle_make/ModelID/vehicle_model/VariantID/VariantName/Fuel/SeatingCapacity

# log("OK", f"Excel masters loaded in {time.time()-t0:.1f}s")

# # ── Company dict ──────────────────────────────────────────────────────────────
# company_by_id = {si(r['company_id']): r for _, r in df_company.iterrows()}

# # ── Segment  (id → name, name → id) ──────────────────────────────────────────
# seg_name_to_id = {str(r['segment']).strip(): si(r['id']) for _, r in df_segment.iterrows()}
# # Policy-type string → canonical segment name
# PT_SEG = {
#     'COMPREHENSIVE': 'Comprehensive', 'COMP': 'Comprehensive',
#     'TP': 'TP Only',  'TP ONLY': 'TP Only', 'THIRD PARTY': 'TP Only',
#     'SAOD': 'SAOD',   'OD': 'SAOD',         'OWN DAMAGE': 'SAOD',
#     'TW NEW': 'Comprehensive',   # new business TW → Comprehensive
# }

# # ── Subproduct  (motor only, name → id) ──────────────────────────────────────
# df_sp_motor = df_subprod[df_subprod.get('product_name', pd.Series(dtype=str)).str.upper() == 'MOTOR'] \
#               if 'product_name' in df_subprod.columns else df_subprod
# subprod_id  = {str(r['sub_product_name']).strip(): si(r['sub_product_id'])
#                for _, r in df_sp_motor.iterrows()}

# # ── Vehicle type  (name → id, sub_product → default id) ──────────────────────
# vt_name_id = {str(r['vehicle_type']).strip(): si(r['id']) for _, r in df_vtype.iterrows()}
# vt_sub_default = {}   # sub_product_name → first vehicle_type_id for that sub
# for _, r in df_vtype.iterrows():
#     sp = str(r['sub_product_name']).strip()
#     if sp not in vt_sub_default:
#         vt_sub_default[sp] = si(r['id'])

# # ── Fuel  (name → id) ─────────────────────────────────────────────────────────
# fuel_id = {str(r['fuel_type']).strip().upper(): si(r['id']) for _, r in df_fuel_xl.iterrows()}
# # Normalise varied fuel strings from MMV sheets → canonical names
# _FUEL_NORM = {
#     'P':'PETROL','PETROL':'PETROL',
#     'D':'DIESEL','DIESEL':'DIESEL',
#     'ELECTRIC':'ELECTRIC','E':'ELECTRIC','B':'ELECTRIC',
#     'CNG':'CNG-LPG','LPG':'CNG-LPG','C':'CNG-LPG','CNG-LPG':'CNG-LPG',
#     'CNG/PETROL':'CNG-LPG','HYBRID':'PETROL','PETROL/ELECTRIC':'ELECTRIC',
# }
# def fuel_lookup(raw):
#     k = _FUEL_NORM.get(str(raw).strip().upper(), '')
#     return fuel_id.get(k, -1), k

# # =============================================================================
# #  MMV INDEX BUILD
# # =============================================================================
# log("INFO", "Building MMV indexes …")
# t1 = time.time()

# # ── TW / Private Car MMV  (global — no company scope) ─────────────────────────
# tw_make  = {}   # make_norm → (make_id, make_name)
# tw_model = {}   # (make_norm, model_norm) → (model_id, model_name)
# tw_var   = {}   # (make_norm, model_norm, var_norm) → (var_id, var_name, fuel_name, seating, cc, geared)
# tw_makes_sorted = []          # sorted list for substring scanning
# tw_models_of_make = {}        # make_norm → set of (model_norm, model_name, model_id)

# for _, m in df_tw_mmv.iterrows():
#     mkn = nt(m.get('vehicle_make',''));  mdn = nt(m.get('vehicle_model',''))
#     vrn = nt(m.get('VariantDisplayName',''))
#     mid = si(m.get('MakeID',-1)); mdid= si(m.get('ModelID',-1)); vid= si(m.get('VariantID',-1))
#     fid_v, fn = fuel_lookup(m.get('Fuel',''))
#     seat= si(m.get('SeatingCapacity',-1)); cc = si(m.get('CC',-1))
#     gear= 1 if str(m.get('Is_Geared_Vehicle','')).upper() in ('TRUE','1','YES') else 0

#     if mkn and mid != -1:
#         tw_make.setdefault(mkn, (mid, str(m.get('vehicle_make','')).strip()))
#     if mkn and mdn and mdid != -1:
#         tw_model.setdefault((mkn, mdn), (mdid, str(m.get('vehicle_model','')).strip()))
#         tw_models_of_make.setdefault(mkn, set()).add((mdn, str(m.get('vehicle_model','')).strip(), mdid))
#     if mkn and mdn and vrn and vid != -1:
#         tw_var.setdefault((mkn, mdn, vrn),
#                           (vid, str(m.get('VariantDisplayName','')).strip(), fn, seat, cc, gear))

# tw_makes_sorted = sorted([(k, tw_make[k][1], tw_make[k][0]) for k in tw_make],
#                          key=lambda x: len(x[0]), reverse=True)

# # ── CV MMV  (scoped by company_code) ──────────────────────────────────────────
# # IMPORTANT: Many makes appear with 2 different MakeIDs (duplicate batches in the master).
# # Strategy: pre-compute the DOMINANT MakeID per (company_code, vehicle_make) = the one
# # with the most variant rows. This is always the authoritative/most-complete batch.
# log("INFO", "Pre-computing dominant MakeID per company+make (resolves duplicate ID issue)...")
# _cv_make_row_count = {}   # (code_nt, make_nt, make_id) → row count
# for _, m in df_cv_mmv.iterrows():
#     code = nt(str(m.get('company_code','')))
#     mkn  = nt(str(m.get('vehicle_make','')))
#     mid  = si(m.get('MakeID', -1))
#     if code and mkn and mid != -1:
#         key = (code, mkn, mid)
#         _cv_make_row_count[key] = _cv_make_row_count.get(key, 0) + 1

# # For each (code, make_norm) pick the MakeID with the highest row count
# _dominant_make_id = {}   # (code, make_norm) → dominant make_id
# _all_make_counts  = {}   # (code, make_norm) → {make_id: count}
# for (code, mkn, mid), cnt in _cv_make_row_count.items():
#     _all_make_counts.setdefault((code, mkn), {})[mid] = cnt
# for (code, mkn), id_counts in _all_make_counts.items():
#     _dominant_make_id[(code, mkn)] = max(id_counts, key=id_counts.get)

# # Similarly pre-compute dominant ModelID per (code, make_norm, model_norm)
# _cv_model_row_count = {}
# for _, m in df_cv_mmv.iterrows():
#     code = nt(str(m.get('company_code','')))
#     mkn  = nt(str(m.get('vehicle_make','')))
#     mdn  = nt(str(m.get('vehicle_model','')))
#     mid  = si(m.get('MakeID', -1))
#     mdid = si(m.get('ModelID', -1))
#     if code and mkn and mdn and mdid != -1:
#         key = (code, mkn, mdn, mdid)
#         _cv_model_row_count[key] = _cv_model_row_count.get(key, 0) + 1

# _dominant_model_id = {}
# _all_model_counts  = {}
# for (code, mkn, mdn, mdid), cnt in _cv_model_row_count.items():
#     _all_model_counts.setdefault((code, mkn, mdn), {})[mdid] = cnt
# for (code, mkn, mdn), id_counts in _all_model_counts.items():
#     _dominant_model_id[(code, mkn, mdn)] = max(id_counts, key=id_counts.get)

# log("OK", f"Dominant IDs computed: {len(_dominant_make_id)} make keys, {len(_dominant_model_id)} model keys")

# cv_make  = {}   # (code, make_norm) → (make_id, make_name)
# cv_model = {}   # (code, make_norm, model_norm) → (model_id, model_name)
# cv_var   = {}   # (code, make_norm, model_norm, var_norm) → (var_id, var_name, fuel_name, seating)
# cv_makes_of_co  = {}   # code → sorted list of (make_norm, make_name, make_id)
# cv_models_of_mk = {}   # (code, make_norm) → set of (model_norm, model_name, model_id)
# cv_models_of_co = {}   # code → {model_norm: [(model_id, model_name, make_id, make_name)]}

# for _, m in df_cv_mmv.iterrows():
#     code= nt(m.get('company_code',''))
#     mkn = nt(m.get('vehicle_make',''));  mdn = nt(m.get('vehicle_model',''))
#     vrn = nt(m.get('VariantName',''))
#     mid = si(m.get('MakeID',-1)); mdid= si(m.get('ModelID',-1)); vid= si(m.get('VariantID',-1))
#     fid_v, fn = fuel_lookup(m.get('Fuel',''))
#     seat= si(m.get('SeatingCapacity',-1))
#     mk_display  = str(m.get('vehicle_make','')).strip()
#     md_display  = str(m.get('vehicle_model','')).strip()

#     if code and mkn and mid != -1:
#         dom_mid = _dominant_make_id.get((code, mkn), mid)
#         # Only index this row's make if mid == dominant, OR entry not yet set
#         if mid == dom_mid or (code, mkn) not in cv_make:
#             cv_make[(code, mkn)] = (dom_mid, mk_display)
#         cv_makes_of_co.setdefault(code, set()).add((mkn, mk_display, dom_mid))

#     if code and mkn and mdn and mdid != -1:
#         dom_mid  = _dominant_make_id.get((code, mkn), mid)
#         dom_mdid = _dominant_model_id.get((code, mkn, mdn), mdid)
#         if mdid == dom_mdid or (code, mkn, mdn) not in cv_model:
#             cv_model[(code, mkn, mdn)] = (dom_mdid, md_display)
#         cv_models_of_mk.setdefault((code, mkn), set()).add((mdn, md_display, dom_mdid))
#         cv_models_of_co.setdefault(code, {}).setdefault(mdn, [])
#         # Only add if not already present with same model_id
#         existing_ids = {x[0] for x in cv_models_of_co[code][mdn]}
#         if dom_mdid not in existing_ids:
#             cv_models_of_co[code][mdn].append((dom_mdid, md_display, dom_mid, mk_display))

#     if code and mkn and mdn and vrn and vid != -1:
#         cv_var.setdefault((code, mkn, mdn, vrn),
#                           (vid, str(m.get('VariantName','')).strip(), fn, seat))

# for k in list(cv_makes_of_co.keys()):
#     cv_makes_of_co[k] = sorted(list(cv_makes_of_co[k]), key=lambda x: len(x[0]), reverse=True)

# log("OK", (f"MMV indexes built in {time.time()-t1:.1f}s  |  "
#            f"TW/PC makes={len(tw_make)} models={len(tw_model)} variants={len(tw_var)}  |  "
#            f"CV make-rows={len(cv_make)} model-rows={len(cv_model)} variant-rows={len(cv_var)}"))

# # =============================================================================
# #  COMPANY SELECTION
# # =============================================================================
# print("\n" + "="*75)
# print("AVAILABLE COMPANIES")
# print("="*75)
# for cid, row in sorted(company_by_id.items()):
#     print(f"  {cid:3d}  |  {str(row['company_code']):20s}  |  {row['company_name']}")
# print("="*75)

# while True:
#     try:
#         CID      = int(input("\nEnter company_id: ").strip())
#         CROW     = company_by_id[CID]
#         CCODE    = str(CROW['company_code']).strip()
#         CCODE_NT = nt(CCODE)
#         log("OK", f"Company → {CROW['company_name']}  (id={CID}  code={CCODE})")
#         break
#     except (ValueError, KeyError):
#         log("ERR", "Invalid company_id, try again.")

# OUT_FILE = os.path.join(OUT_DIR, f"{CCODE}-Payin-Config.xlsx")

# _api_key = os.getenv("OPENAI_API_KEY","").strip()
# _model   = os.getenv("OPENAI_MODEL","gpt-4.1-mini")
# if _api_key: log("OK",   f"OpenAI key found — model: {_model}")
# else:        log("WARN", "No OPENAI_API_KEY — heuristic-only remark parsing")

# # =============================================================================
# #  PURE-PARSING HELPERS
# # =============================================================================

# def norm_pt(v):
#     """Policy type raw string → COMP / TP / SAOD."""
#     s = str(v).strip().upper()
#     if s in ('COMP','COMPREHENSIVE'):        return 'COMP'
#     if s in ('TP','TP ONLY','THIRD PARTY'):  return 'TP'
#     if s in ('SAOD','OD','OWN DAMAGE'):      return 'SAOD'
#     if s == 'TW NEW':                        return 'COMP'
#     return 'COMP'

# def vehicle_info(seg_text):
#     """SEGMENT column → (sub_product_name, sub_product_id, default_vehicle_type_id)."""
#     s = str(seg_text).strip().upper()
#     # Two Wheeler
#     if re.search(r'\bTW\b', s) or re.search(r'\b2W\b', s):
#         sp = 'Two Wheeler'
#         return sp, subprod_id.get(sp, -1), vt_name_id.get('TW Bike', vt_sub_default.get(sp, -1))
#     # Private Car
#     if re.search(r'\bPVT\s*CAR\b', s) or re.search(r'\bPRIVATE\s*CAR\b', s) or re.search(r'\bPC\b', s):
#         sp = 'Private Car'
#         return sp, subprod_id.get(sp, -1), vt_name_id.get('Private Car', vt_sub_default.get(sp, -1))
#     # Passenger Vehicle
#     if re.search(r'\bPCV\b', s) or re.search(r'\bPASSENGER\b', s):
#         sp = 'Passenger Vehicle'
#         return sp, subprod_id.get(sp, -1), vt_name_id.get('Auto rikshaw', vt_sub_default.get(sp, -1))
#     # Miscellaneous — Tractor / Harvester
#     if 'TRACTOR' in s or 'HARVESTER' in s:
#         sp = 'Miscellaneous Vehicle'
#         return sp, subprod_id.get(sp, -1), vt_name_id.get('Agriculture Tractor', vt_sub_default.get(sp, -1))
#     # Miscellaneous — MISD / MISC
#     if 'MISD' in s or 'MISC' in s:
#         sp = 'Miscellaneous Vehicle'
#         return sp, subprod_id.get(sp, -1), vt_name_id.get('Non Tractor', vt_sub_default.get(sp, -1))
#     # Goods Vehicle — GCV / GVW / CV
#     sp = 'Goods Vehicle'
#     return sp, subprod_id.get(sp, -1), vt_name_id.get('Truck', vt_sub_default.get(sp, -1))

# def parse_age(s):
#     s = str(s).strip()
#     if s.lower() in ('', 'nan', 'none', 'new', 'n/a'): return 0, 700
#     su = s.upper()
#     m = re.search(r'UPTO\s+(\d+)\s*(YEAR|YR|MONTH|MTH|MO)', su)
#     if m: n, u = int(m.group(1)), m.group(2); return 0, n*12 if u.startswith('Y') else n
#     m = re.search(r'(\d+)\s*TO\s*(\d+)\s*(YEAR|YR)?', su)
#     if m: return int(m.group(1))*12, int(m.group(2))*12
#     m = re.match(r'>\s*(\d+)\s*[-]\s*(\d+)(\+?)\s*[Yy]', s)
#     if m: return int(m.group(1))*12+1, (700 if m.group(3)=='+' else int(m.group(2))*12)
#     m = re.match(r'>\s*(\d+)\+?\s*[Yy]', s)
#     if m: return int(m.group(1))*12+1, 700
#     m = re.match(r'^(\d+)$', s.strip())
#     if m: return 0, int(m.group(1))*12
#     return 0, 700

# def parse_cc(s):
#     if not s or str(s).strip().lower() in ('', 'nan', 'none'): return 0, 99999, -1
#     s = str(s).strip().upper().replace('CC','').strip()
#     for pat, fn in [
#         (r'^<\s*(\d+)$',                 lambda m: (0, int(m.group(1))-1, 1)),
#         (r'^>\s*(\d+)$',                 lambda m: (int(m.group(1))+1, 99999, 1)),
#         (r'^>\s*(\d+)\s*[-]\s*(\d+)$',  lambda m: (int(m.group(1))+1, int(m.group(2)), 1)),
#         (r'^>=\s*(\d+)\s*[-]\s*(\d+)$', lambda m: (int(m.group(1)), int(m.group(2)), 1)),
#         (r'^(\d+)\s*[-]\s*(\d+)$',      lambda m: (int(m.group(1)), int(m.group(2)), 1)),
#         (r'^(\d+)$',                     lambda m: (int(m.group(1)), int(m.group(1)), 1)),
#     ]:
#         mt = re.match(pat, s)
#         if mt: return fn(mt)
#     return 0, 99999, -1

# def parse_idv(text):
#     su = str(text).upper()
#     m = re.search(r'IDV\s+(?:UPTO|UP\s*TO|<=?)\s*([\d.]+)\s*(?:LAC|LAKH|L\b)', su)
#     if m: return 1, 0.0, float(m.group(1))
#     m = re.search(r'IDV\s+([\d.]+)\s*[-]\s*([\d.]+)\s*(?:LAC|LAKH)', su)
#     if m: return 1, float(m.group(1)), float(m.group(2))
#     return -1, 0.0, 0.0

# def parse_weight(text):
#     su = str(text).upper()
#     m = re.search(r'([\d.]+)\s*KG\b', su)
#     if m: return 1, 0, int(float(m.group(1)))
#     m = re.search(r'UPTO\s+([\d.]+)\s*T(?:ON|ONNE)?', su)
#     if m: return 1, 0, int(float(m.group(1))*1000)
#     m = re.search(r'>\s*([\d.]+)\s*[-]\s*([\d.]+)\s*T(?:ON|ONNE)?', su)
#     if m: return 1, int(float(m.group(1))*1000)+1, int(float(m.group(2))*1000)
#     m = re.search(r'\b([\d.]+)\s*T(?:ON|ONNE)?\b', su)
#     if m: return 1, 0, int(float(m.group(1))*1000)
#     m = re.search(r'GVW\s+([\d.]+)', su)
#     if m: v = float(m.group(1)); return 1, 0, int(v*1000) if v < 200 else int(v)
#     return -1, 0, 99999

# def parse_seating(text):
#     su = str(text).upper()
#     m = re.search(r'(?:SC|SEATING\s*CAPACITY|SEATING)\s*UPTO\s*(\d+)', su)
#     if m: return 1, 1, int(m.group(1))
#     m = re.search(r'(?:SC|SEATING\s*CAPACITY|SEATING)\s*(\d+)\s*[-]\s*(\d+)', su)
#     if m: return 1, int(m.group(1)), int(m.group(2))
#     m = re.search(r'(?:SC|SEATING\s*CAPACITY|SEATING)\s*(\d+)', su)
#     if m: v = int(m.group(1)); return 1, v, v
#     m = re.search(r'(\d+)\s*SEATER', su)
#     if m: v = int(m.group(1)); return 1, v, v
#     return -1, -1, -1

# def parse_fuel_text(text):
#     su = str(text).upper()
#     if 'ELECTRIC' in su or ' EV ' in su: return fuel_id.get('ELECTRIC',-1), 'ELECTRIC'
#     if 'CNG' in su or 'LPG' in su:       return fuel_id.get('CNG-LPG',-1),  'CNG-LPG'
#     if 'DIESEL' in su:                   return fuel_id.get('DIESEL',-1),    'DIESEL'
#     if 'PETROL' in su:                   return fuel_id.get('PETROL',-1),    'PETROL'
#     return -1, ''

# def ncb_flag(text):
#     su = str(text).upper()
#     if any(x in su for x in ('WITHOUT NCB','W/O NCB','NON NCB','NON-NCB','ZERO NCB')): return 0
#     if 'NCB' in su: return 1
#     return -1

# def irda_flag(text):
#     su = str(text).upper()
#     return 1 if any(x in su for x in ('IRDA TP','IRDA RATE','IRDA')) else -1

# def cpa_flag(text):
#     su = str(text).upper()
#     if 'CPA' in su and any(x in su for x in ('INCLUD','WITH CPA')): return 1
#     return -1

# def keep_included(text):
#     """Return only the 'included' portion — strip excluded/except/rejected chunks."""
#     s = str(text).strip()
#     if not s or s.lower() in ('nan', 'none'): return ''
#     kept = []
#     for chunk in re.split(r';', s):
#         up = chunk.upper()
#         if any(t in up for t in ('DECLIN','REJECT')): continue
#         for tok in (' BUT ',' EXCEPT ',' EXCLUDE ',' OTHER THAN ',' NOT CONSIDER'):
#             idx = up.find(tok)
#             if idx != -1: chunk = chunk[:idx]; break
#         chunk = chunk.strip()
#         if chunk: kept.append(chunk)
#     return ' '.join(x for x in kept if x)

# def is_pure_exclusion(text):
#     """True if the remark has ONLY exclusion language and no inclusion items."""
#     s = str(text).strip().upper()
#     if not s or s in ('NAN','NONE'): return False
#     EXCL = ('EXCEPT','EXCLUDE','OTHER THAN','REJECT','DECLIN','NOT CONSIDER',
#             'HR 68','EXCLUDED')
#     INCL = ('ONLY','INCLUD','HONDA','BAJAJ','TATA','MARUTI','HYUNDAI','KIA',
#             'MAHINDRA','IDV','NCB','ZONE','BRANCH')
#     if any(t in s for t in INCL): return False
#     return any(t in s for t in EXCL)

# # =============================================================================
# #  MMV RESOLUTION
# # =============================================================================

# def resolve_tw(raw_make, raw_model, raw_variant, remark):
#     mkn = nt(raw_make) if raw_make else ''
#     mdn = nt(raw_model) if raw_model else ''
#     vrn = nt(raw_variant) if raw_variant else ''
#     mk_id=-1; mk_name=''; md_id=-1; md_name=''; vr_id=-1; vr_name=''
#     fuel=''; seat=-1; cc=-1; gear=-1
#     inc = nt(keep_included(remark))

#     # 1. Make
#     if mkn:
#         if mkn in tw_make:
#             mk_id, mk_name = tw_make[mkn]
#         else:
#             mw = set(mkn.split())
#             for cn, cname, cid in tw_makes_sorted:
#                 if mkn in cn or cn in mkn or mw.issubset(set(cn.split())):
#                     mk_id, mk_name, mkn = cid, cname, cn; break
#         if mk_id == -1:  # scan remark
#             for cn, cname, cid in tw_makes_sorted:
#                 if cn and re.search(rf'\b{re.escape(cn)}\b', inc):
#                     mk_id, mk_name, mkn = cid, cname, cn; break

#     # 2. Model (only if mentioned)
#     if mdn and mk_id != -1:
#         if (mkn, mdn) in tw_model:
#             md_id, md_name = tw_model[(mkn, mdn)]
#         else:
#             for cmn, cmname, cmid in sorted(tw_models_of_make.get(mkn, set()),
#                                             key=lambda x: len(x[0]), reverse=True):
#                 if mdn in cmn or cmn in mdn:
#                     md_id, md_name, mdn = cmid, cmname, cmn; break

#     # 3. Variant (only if mentioned)
#     if vrn and mk_id != -1 and md_id != -1:
#         key = (mkn, mdn, vrn)
#         if key in tw_var:
#             vr_id, vr_name, fuel, seat, cc, gear = tw_var[key]
#         else:
#             for (cmn,cmdn,cvn), (vid,vname,vf,vs,vc,vg) in tw_var.items():
#                 if cmn==mkn and cmdn==mdn and (vrn in cvn or cvn in vrn):
#                     vr_id,vr_name,fuel,seat,cc,gear = vid,vname,vf,vs,vc,vg; break

#     # Infer fuel/seating from first matching variant when no variant given
#     if mk_id != -1 and md_id != -1 and not vrn:
#         for (cmn,cmdn,_),(vid,vname,vf,vs,vc,vg) in tw_var.items():
#             if cmn==mkn and cmdn==mdn and vf:
#                 fuel,seat,cc,gear = vf,vs,vc,vg; break

#     return mk_id, mk_name, md_id, md_name, vr_id, vr_name, fuel, seat, cc, gear


# def resolve_cv(raw_make, raw_model, raw_variant, remark):
#     code = CCODE_NT
#     mkn  = nt(raw_make)    if raw_make    else ''
#     mdn  = nt(raw_model)   if raw_model   else ''
#     vrn  = nt(raw_variant) if raw_variant else ''
#     mk_id=-1; mk_name=''; md_id=-1; md_name=''; vr_id=-1; vr_name=''
#     fuel=''; seat=-1
#     inc = nt(keep_included(remark))

#     # 1. Make
#     if mkn:
#         if (code, mkn) in cv_make:
#             mk_id, mk_name = cv_make[(code, mkn)]
#         else:
#             mw = set(mkn.split())
#             for cn, cname, cid in cv_makes_of_co.get(code, []):
#                 if mkn in cn or cn in mkn or mw.issubset(set(cn.split())):
#                     mk_id, mk_name, mkn = cid, cname, cn; break
#         if mk_id == -1:  # scan remark
#             for cn, cname, cid in cv_makes_of_co.get(code, []):
#                 if cn and re.search(rf'\b{re.escape(cn)}\b', inc):
#                     mk_id, mk_name, mkn = cid, cname, cn; break

#     # 2. Model
#     if mdn:
#         if mk_id != -1 and (code, mkn, mdn) in cv_model:
#             md_id, md_name = cv_model[(code, mkn, mdn)]
#         elif mk_id != -1:
#             for cmn,cmname,cmid in sorted(cv_models_of_mk.get((code, mkn), set()),
#                                           key=lambda x: len(x[0]), reverse=True):
#                 if mdn in cmn or cmn in mdn:
#                     md_id, md_name, mdn = cmid, cmname, cmn; break
#         if md_id == -1:
#             hits = cv_models_of_co.get(code, {}).get(mdn, [])
#             if hits: md_id, md_name, mk_id, mk_name = hits[0]

#     # 3. Variant
#     if vrn and mk_id != -1 and md_id != -1:
#         key = (code, mkn, mdn, vrn)
#         if key in cv_var:
#             vr_id, vr_name, fuel, seat = cv_var[key]
#         else:
#             for (ck,cmn,cmdn,cvn),(vid,vname,vf,vs) in cv_var.items():
#                 if ck==code and cmn==mkn and cmdn==mdn and (vrn in cvn or cvn in vrn):
#                     vr_id,vr_name,fuel,seat = vid,vname,vf,vs; break

#     if mk_id != -1 and md_id != -1 and not vrn:
#         for (ck,cmn,cmdn,_),(vid,vname,vf,vs) in cv_var.items():
#             if ck==code and cmn==mkn and cmdn==mdn and vf:
#                 fuel, seat = vf, vs; break

#     return mk_id, mk_name, md_id, md_name, vr_id, vr_name, fuel, seat

# # =============================================================================
# #  OPENAI REMARK PARSER
# # =============================================================================

# def heuristic(remark, seg, pt):
#     su = str(remark).upper()
#     fid,fn  = parse_fuel_text(remark)
#     iw,fw,tw_= parse_weight(remark)
#     iss,fs,ts= parse_seating(remark)
#     ii,fi,ti = parse_idv(remark)
#     return {
#         'vehicle_make':'','vehicle_model':'','vehicle_variant':'',
#         'is_with_ncb':ncb_flag(su), 'is_irda_tp':irda_flag(su),
#         'is_cpa_included':cpa_flag(su),
#         'fuel_type':fn,
#         'seating_cap':-1,'from_seating':fs,'to_seating':ts,
#         'is_weight':iw,'from_weight_kg':fw,'to_weight_kg':tw_,
#         'idv_cap':ii,'from_idv':fi,'to_idv':ti,
#         'is_cc':-1,'from_cc':0,'to_cc':99999,
#     }

# _nc=0; _nh=0; _ne=0; _ms=0

# def parse_remark(remark, co_name, seg, pt, row_n=0):
#     global _nc, _ne, _ms
#     ak = os.getenv("OPENAI_API_KEY","").strip()
#     if not ak: return heuristic(remark, seg, pt)

#     included = keep_included(remark)
#     short    = (remark[:80]+'…') if len(remark)>80 else remark
#     log("API", f"Row {row_n:>4} | call #{_nc+1} | seg={seg!r:20} | {short!r}")

#     prompt = f"""You are an expert Indian motor insurance data extractor.
# Analyse the remark carefully. Return ONLY a valid JSON object — no markdown, no preamble — with EXACTLY these keys:

#   vehicle_make    : INCLUDED makes comma-separated (e.g. "HONDA,HYUNDAI,KIA"). Empty if none.
#   vehicle_model   : INCLUDED models comma-separated. Empty if none.
#   vehicle_variant : Specific variant name if mentioned, else "".
#   is_with_ncb     : 1 if NCB cases INCLUDED. 0 if WITHOUT/NON/ZERO NCB. -1 if not mentioned.
#   is_irda_tp      : 1 if IRDA TP rate mentioned, -1 otherwise.
#   is_cpa_included : 1 if CPA included/mentioned, -1 otherwise.
#   fuel_type       : DIESEL | PETROL | ELECTRIC | CNG-LPG | "" (empty if not mentioned).
#   seating_cap     : exact integer if single seating value given, -1 otherwise.
#   from_seating    : lower seating bound (int), -1 if N/A.
#   to_seating      : upper seating bound (int), -1 if N/A.
#   is_weight       : 1 if GVW/weight/tonnage mentioned, -1 otherwise.
#   from_weight_kg  : lower weight KG int, 0 if N/A.
#   to_weight_kg    : upper weight KG int, 99999 if N/A.
#   idv_cap         : 1 if IDV cap mentioned, -1 otherwise.
#   from_idv        : lower IDV in Lacs (float), 0 if N/A.
#   to_idv          : upper IDV in Lacs (float), 0 if N/A.
#   is_cc           : 1 if engine CC/capacity mentioned, -1 otherwise.
#   from_cc         : lower CC int, 0 if N/A.
#   to_cc           : upper CC int, 99999 if N/A.

# RULES — READ CAREFULLY:
# 1. ONLY list makes/models that are INCLUDED. IGNORE anything after EXCEPT / BUT / EXCLUDE /
#    OTHER THAN / REJECT / DECLINE / ONLY EXCEPT.
# 2. If remark has ONLY exclusion text (e.g. "Except TATA" or "HR 68 EXCLUDED"), set
#    vehicle_make="" and vehicle_model="".
# 3. SC = Seating Capacity. "SC 7" → seating_cap=7. "SC upto 7" → from_seating=1, to_seating=7.
# 4. Tons → KG: 1 ton = 1000 KG. "7.5T" → to_weight_kg=7500.
# 5. "IDV upto 10 lacs" → idv_cap=1, from_idv=0, to_idv=10.
# 6. "Upto 1500 CC" → is_cc=1, from_cc=0, to_cc=1500.
# 7. TRACTOR segment → vehicle_variant="Agriculture Tractor".

# company: {co_name}
# segment: {seg}
# policy_type: {pt}
# remark_original: {remark}
# remark_included_only: {included}
# """
#     body = {"model": _model,
#             "messages": [
#                 {"role":"system","content":"Return only valid JSON, no markdown, no explanation."},
#                 {"role":"user","content":prompt}],
#             "response_format":{"type":"json_object"}, "temperature":0}
#     req = urllib.request.Request(
#         "https://api.openai.com/v1/chat/completions",
#         data=json.dumps(body).encode(),
#         headers={"Content-Type":"application/json","Authorization":f"Bearer {ak}"},
#         method="POST")
#     ts = time.time()
#     try:
#         with urllib.request.urlopen(req, timeout=30) as r:
#             p = json.loads(json.loads(r.read())["choices"][0]["message"]["content"])
#         ms = int((time.time()-ts)*1000); _nc+=1; _ms+=ms
#         def _s(k,d=""): return str(p.get(k,d)).strip()
#         def _i(k,d=-1):
#             try: return int(p.get(k,d))
#             except: return d
#         def _f(k,d=0.0):
#             try: return float(p.get(k,d))
#             except: return d
#         result = {
#             'vehicle_make':_s('vehicle_make'),'vehicle_model':_s('vehicle_model'),
#             'vehicle_variant':_s('vehicle_variant'),
#             'is_with_ncb':_i('is_with_ncb'),'is_irda_tp':_i('is_irda_tp'),
#             'is_cpa_included':_i('is_cpa_included'),
#             'fuel_type':_s('fuel_type').upper(),
#             'seating_cap':_i('seating_cap'),'from_seating':_i('from_seating'),
#             'to_seating':_i('to_seating'),
#             'is_weight':_i('is_weight'),
#             'from_weight_kg':_i('from_weight_kg',0),'to_weight_kg':_i('to_weight_kg',99999),
#             'idv_cap':_i('idv_cap'),'from_idv':_f('from_idv'),'to_idv':_f('to_idv'),
#             'is_cc':_i('is_cc'),'from_cc':_i('from_cc',0),'to_cc':_i('to_cc',99999),
#         }
#         log("OK", (f"Row {row_n:>4} | {ms}ms | "
#                    f"make={result['vehicle_make']!r:15} model={result['vehicle_model']!r:12} "
#                    f"ncb={result['is_with_ncb']} irda={result['is_irda_tp']} "
#                    f"cc={result['is_cc']} fuel={result['fuel_type']!r}"))
#         return result
#     except urllib.error.HTTPError as e:
#         _ne+=1; log("ERR", f"Row {row_n:>4} | HTTP {e.code} → heuristic")
#     except Exception as e:
#         _ne+=1; log("ERR", f"Row {row_n:>4} | {e} → heuristic")
#     return heuristic(remark, seg, pt)

# # =============================================================================
# #  OUTPUT COLUMN  ORDER
# # =============================================================================
# COLS = [
#     'id','company_id','company_code','segment_id','segment',
#     'subproduct_id','sub_product_name','lob_id','lob_name',
#     'business_type_id','business_type','is_highend_lob',
#     'rto_group_id','rto_group_name',
#     'payin_od_rate','payin_tp_rate','payout_od_rate','payout_tp_rate',
#     'extra_tp_rate','eff_from_date','eff_to_date',
#     'fuel_type_id','fuel_type',
#     'is_on_net','is_one_year_pay_on_newbusiness','is_cpa_included',
#     'is_geared_vehicle','is_cc_considered','from_cc','to_cc',
#     'is_premium_considered','from_premium','to_premium',
#     'is_mmv_considered','make_id','vehicle_make','model_id','vehicle_model',
#     'variant_id','vehicle_variant',
#     'is_seating_cap_consider','from_seating_cap','to_seating_cap',
#     'is_no_of_wheel_consider','from_no_of_wheel','to_no_of_wheel',
#     'vehicle_type_id','ppi_in','ppi_out',
#     'is_irda_tp_included','is_longterm_renewal_pay',
#     'is_weightage_considered','from_weightage_kg','to_weightage_kg',
#     'is_nil_dep_considered','is_organization_type',
#     'from_age_month','to_age_month',
#     'is_with_ncb','is_idv_cap_consider','from_idv','to_idv',
#     'is_breakin_consider','is_active',
# ]

# def make_row(
#     sid,sname,spid,spname,rto_id,rto_name,
#     pod,ptp,pod2,ptp2,
#     fid,fname,is_on_net,cpa,geared,
#     iscc,fcc,tcc,
#     iswt,fwt,twt,
#     vtid,
#     is_mmv,mkid,mkname,mdid,mdname,vrid,vrname,
#     issc,fsc,tsc,
#     nw,fnw,tnw,
#     fage,tage,
#     ncb,irda,
#     idvcap,fidv,tidv,
# ):
#     return {
#         'id':0, 'company_id':CID, 'company_code':CCODE,
#         'segment_id':sid, 'segment':sname,
#         'subproduct_id':spid, 'sub_product_name':spname,
#         'lob_id':-1, 'lob_name':'',
#         'business_type_id':-1, 'business_type':'Not Considered', 'is_highend_lob':False,
#         'rto_group_id':rto_id, 'rto_group_name':rto_name,
#         'payin_od_rate':pod, 'payin_tp_rate':ptp,
#         'payout_od_rate':pod2, 'payout_tp_rate':ptp2,
#         'extra_tp_rate':0, 'eff_from_date':'2026-01-01', 'eff_to_date':'2026-01-16',
#         'fuel_type_id':fid, 'fuel_type':fname,
#         'is_on_net':is_on_net,
#         'is_one_year_pay_on_newbusiness':-1, 'is_cpa_included':cpa,
#         'is_geared_vehicle':geared,
#         'is_cc_considered':iscc, 'from_cc':fcc, 'to_cc':tcc,
#         'is_premium_considered':-1, 'from_premium':-1, 'to_premium':-1,
#         'is_mmv_considered':is_mmv,
#         'make_id':mkid, 'vehicle_make':mkname,
#         'model_id':mdid, 'vehicle_model':mdname,
#         'variant_id':vrid, 'vehicle_variant':vrname,
#         'is_seating_cap_consider':issc, 'from_seating_cap':fsc, 'to_seating_cap':tsc,
#         'is_no_of_wheel_consider':nw, 'from_no_of_wheel':fnw, 'to_no_of_wheel':tnw,
#         'vehicle_type_id':vtid, 'ppi_in':0, 'ppi_out':0,
#         'is_irda_tp_included':irda, 'is_longterm_renewal_pay':-1,
#         'is_weightage_considered':iswt, 'from_weightage_kg':fwt, 'to_weightage_kg':twt,
#         'is_nil_dep_considered':-1, 'is_organization_type':-1,
#         'from_age_month':fage, 'to_age_month':tage,
#         'is_with_ncb':ncb,
#         'is_idv_cap_consider':idvcap, 'from_idv':fidv, 'to_idv':tidv,
#         'is_breakin_consider':-1, 'is_active':True,
#     }

# # =============================================================================
# #  PROCESS ONE INPUT FILE
# # =============================================================================

# def process(input_file):
#     global _nh
#     log("INFO", f"Reading: {input_file}")
#     tr = time.time()
#     df = pd.read_excel(input_file)
#     df.columns = [c.strip() for c in df.columns]
#     log("OK", f"Loaded {len(df)} rows in {time.time()-tr:.1f}s | cols: {list(df.columns)}")

#     def col(*names):
#         for n in names:
#             if n in df.columns: return n
#         return None

#     c_seg  = col('SEGMENT','Segment')
#     c_pt   = col('POLICY TYPE','Policy Type','POLICYTYPE')
#     c_loc  = col('LOCATION','Location')
#     c_pay  = col('PAYIN','Payin')
#     c_pout = col('PAYOUT','Payout','Calculated Payout')
#     c_rem  = col('REMARK','Remark','REMARKS','CALCULATION EXPLANATION')
#     c_age  = col('AGE','Age','AGE BAND')
#     c_cc   = col('CC BAND','CC Band','CC')
#     c_tw   = col('TW TYPE','TW Type')
#     c_co   = col('COMPANY NAME','Company Name','COMPANY')

#     log("INFO", f"col-map → seg={c_seg!r} pt={c_pt!r} loc={c_loc!r} "
#         f"pay={c_pay!r} pout={c_pout!r} rem={c_rem!r} age={c_age!r}")

#     # estimate API calls
#     uniq_rem = set()
#     if c_rem:
#         for v in df[c_rem].fillna('').astype(str): uniq_rem.add(v.strip())
#     log("INFO", f"Unique remarks: {len(uniq_rem)} → "
#         + (f"~{len(uniq_rem)} API calls" if _api_key else "heuristic only"))

#     out_rows = []; cache = {}; total = len(df); tp = time.time()
#     print(f"\n  {'='*58}\n  Processing {total} rows …\n  {'='*58}\n")

#     for idx, (_, row) in enumerate(df.iterrows(), 1):
#         bar(idx, total)

#         if idx == 1 or idx % 25 == 0 or idx == total:
#             el = time.time()-tp; rate = idx/el if el else 0; eta = (total-idx)/rate if rate else 0
#             avg = (_ms/_nc) if _nc else 0
#             log("INFO", (f"Row {idx:>4}/{total} | {el:.0f}s | ETA {eta:.0f}s | "
#                          f"rate={rate:.1f}/s | api={_nc} avg={avg:.0f}ms | "
#                          f"cache={_nh} err={_ne} | out={len(out_rows)}"))

#         def g(c, d=''):
#             if c is None: return d
#             v = row.get(c, d)
#             return v if v is not None and str(v).strip() not in ('nan','None','NaN') else d

#         # ── Policy type & segment ─────────────────────────────────────────────
#         pt_raw = str(g(c_pt, 'COMP')).strip()
#         pt     = norm_pt(pt_raw)
#         seg_nm = PT_SEG.get(pt, 'Comprehensive')
#         seg_id = seg_name_to_id.get(seg_nm, 1)

#         # ── Rates ─────────────────────────────────────────────────────────────
#         payin  = sf(g(c_pay, 0))
#         payout = sf(g(c_pout, 0))
#         if   pt == 'TP':   pod, ptp, pod2, ptp2 = 0, payin, 0, payout
#         elif pt == 'SAOD': pod, ptp, pod2, ptp2 = payin, 0, payout, 0
#         else:              pod = ptp = payin; pod2 = ptp2 = payout

#         # ── Vehicle / subproduct ──────────────────────────────────────────────
#         seg_text = str(g(c_seg, '')).strip()
#         spname, spid, vtid = vehicle_info(seg_text)
#         is_cv   = spname in ('Goods Vehicle','Passenger Vehicle','Miscellaneous Vehicle')

#         # ── Location (rto_group_id always 0 per corrections.txt) ─────────────
#         rto_name = str(g(c_loc, '')).strip()
#         rto_id   = 0

#         # ── Other columns ─────────────────────────────────────────────────────
#         remark   = str(g(c_rem, '')).strip()
#         co_name  = str(g(c_co, CCODE)).strip()
#         fage, tage = parse_age(str(g(c_age, '')))
#         fcc0, tcc0, iscc0 = parse_cc(str(g(c_cc, '')))

#         # TW geared / vehicle-type override
#         tw_raw = str(g(c_tw, '')).strip().lower()
#         geared = -1
#         if spname == 'Two Wheeler' and tw_raw:
#             if 'scooter'  in tw_raw: vtid = vt_name_id.get('TW Scooter', vtid); geared = 0
#             elif 'bike'   in tw_raw: vtid = vt_name_id.get('TW Bike',    vtid); geared = 1
#             elif 'electric' in tw_raw: vtid = vt_name_id.get('TW Electric Bike', vtid); geared = -1

#         # ── is_on_net ─────────────────────────────────────────────────────────
#         is_on_net = True if pt == 'COMP' else False

#         # ── OpenAI / cache ────────────────────────────────────────────────────
#         ck = (remark, co_name, seg_text, pt)
#         if ck in cache:
#             _nh += 1
#             log("CACHE", f"Row {idx:>4} | HIT #{_nh} | {remark[:50]!r}")
#             meta = cache[ck]
#         else:
#             meta  = parse_remark(remark, co_name, seg_text, pt, idx)
#             cache[ck] = meta

#         # ── Pull meta fields ──────────────────────────────────────────────────
#         ncb    = meta['is_with_ncb']
#         irda   = meta['is_irda_tp']
#         cpa    = meta['is_cpa_included']
#         raw_mk = meta['vehicle_make']
#         raw_md = meta['vehicle_model']
#         raw_vr = meta['vehicle_variant']

#         # Fuel
#         m_fuel = meta.get('fuel_type', '')
#         if m_fuel and m_fuel in fuel_id: fid_v, fn_v = fuel_id[m_fuel], m_fuel
#         else: fid_v, fn_v = parse_fuel_text(remark)

#         # Seating
#         msc  = meta.get('seating_cap',-1)
#         mfsc = meta.get('from_seating',-1)
#         mtsc = meta.get('to_seating',-1)
#         if msc != -1:                        issc_v,fsc_v,tsc_v = 1,msc,msc
#         elif mfsc != -1 or mtsc != -1:       issc_v=1; fsc_v=mfsc if mfsc!=-1 else 1; tsc_v=mtsc if mtsc!=-1 else 99
#         else:                                issc_v,fsc_v,tsc_v = parse_seating(remark)

#         # Weight
#         m_wt  = meta.get('is_weight',-1)
#         if m_wt == 1: iswt_v,fwt_v,twt_v = 1,meta['from_weight_kg'],meta['to_weight_kg']
#         else:         iswt_v,fwt_v,twt_v = parse_weight(remark)

#         # IDV
#         m_idv = meta.get('idv_cap',-1)
#         if m_idv == 1: idv_v,fidv_v,tidv_v = 1,meta['from_idv'],meta['to_idv']
#         else:          idv_v,fidv_v,tidv_v = parse_idv(remark)

#         # CC  (column-level overrides OpenAI if column exists)
#         iscc_v, fcc_v, tcc_v = iscc0, fcc0, tcc0
#         if iscc_v == -1 and meta.get('is_cc',-1) == 1:
#             iscc_v, fcc_v, tcc_v = 1, meta['from_cc'], meta['to_cc']

#         # Pure exclusion → clear MMV
#         if is_pure_exclusion(remark):
#             raw_mk = raw_md = raw_vr = ''
#             log("INFO", f"Row {idx:>4} | pure-exclusion remark → MMV cleared")

#         # ── Expand multiple makes ─────────────────────────────────────────────
#         make_list = [m.strip() for m in re.split(r'[,&]+', raw_mk) if m.strip()] or ['']
#         if len(make_list) > 1:
#             log("INFO", f"Row {idx:>4} | expanding {len(make_list)} makes: {make_list}")

#         for one_make in make_list:
#             if is_cv:
#                 mkid,mkname,mdid,mdname,vrid,vrname,i_fuel,i_seat = \
#                     resolve_cv(one_make, raw_md, raw_vr, remark)
#                 i_cc=-1; i_gear=-1
#             else:
#                 mkid,mkname,mdid,mdname,vrid,vrname,i_fuel,i_seat,i_cc,i_gear = \
#                     resolve_tw(one_make, raw_md, raw_vr, remark)

#             if one_make and mkid == -1:
#                 log("WARN", f"Row {idx:>4} | make '{one_make}' NOT in MMV for {CCODE}")

#             # Tractor → vehicle_variant override
#             if spname == 'Miscellaneous Vehicle' and 'TRACTOR' in seg_text.upper():
#                 if not vrname: vrname = 'Agriculture Tractor'

#             is_mmv = 1 if (mkid!=-1 or mdid!=-1 or vrid!=-1 or
#                            one_make or raw_md or raw_vr) else -1

#             # Fuel fallback from MMV variant data
#             fin_fid, fin_fn = fid_v, fn_v
#             if fin_fid == -1 and i_fuel and i_fuel in fuel_id:
#                 fin_fid = fuel_id[i_fuel]; fin_fn = i_fuel

#             # Seating fallback from MMV
#             fin_issc, fin_fsc, fin_tsc = issc_v, fsc_v, tsc_v
#             if fin_issc == -1 and i_seat > 0:
#                 fin_issc=1; fin_fsc=i_seat; fin_tsc=i_seat

#             # Geared fallback from MMV variant
#             fin_gear = geared
#             if spname == 'Two Wheeler' and fin_gear == -1 and i_gear != -1:
#                 fin_gear = i_gear

#             out_rows.append(make_row(
#                 seg_id, seg_nm, spid, spname, rto_id, rto_name,
#                 pod, ptp, pod2, ptp2,
#                 fin_fid, fin_fn, is_on_net, cpa, fin_gear,
#                 iscc_v, fcc_v, tcc_v,
#                 iswt_v, fwt_v, twt_v,
#                 vtid,
#                 is_mmv,
#                 mkid, mkname if mkname else one_make,
#                 mdid, mdname if mdname else raw_md,
#                 vrid, vrname if vrname else raw_vr,
#                 fin_issc, fin_fsc, fin_tsc,
#                 -1, -1, -1,
#                 fage, tage,
#                 ncb, irda,
#                 idv_v, fidv_v, tidv_v,
#             ))

#     el = time.time()-tp; avg = (_ms/_nc) if _nc else 0
#     print()
#     log("OK","="*60)
#     log("OK",f"DONE  input={total}  output={len(out_rows)}  time={el:.1f}s ({el/60:.1f}min)")
#     log("OK",f"  API calls={_nc}  avg={avg:.0f}ms  cache={_nh}  errors={_ne}")
#     log("OK","="*60)
#     return pd.DataFrame(out_rows)

# # =============================================================================
# #  MAIN LOOP  — process one or more files, all appended to same output
# # =============================================================================
# input_file = input("\nEnter Shriram input Excel file path: ").strip().strip('"')

# while True:
#     try:
#         df_out = process(input_file)
#         df_out = df_out[[c for c in COLS if c in df_out.columns]]

#         if os.path.exists(OUT_FILE):
#             log("INFO", f"Appending to existing: {OUT_FILE}")
#             df_out = pd.concat([pd.read_excel(OUT_FILE), df_out], ignore_index=True)

#         df_out.to_excel(OUT_FILE, index=False)
#         log("OK", f"Saved → {OUT_FILE}  ({len(df_out)} rows)")
#         log("OK", f"Log   → {_log_path}")

#     except Exception as e:
#         import traceback
#         log("ERR", f"FATAL: {e}")
#         traceback.print_exc()

#     print("\n" + "="*75)
#     print("  1  Process another Shriram file (appends to same output)")
#     print("  2  Exit")
#     ch = input("Choice: ").strip()
#     if ch == "2":
#         log("OK","Goodbye!")
#         if _log_file: _log_file.close()
#         break
#     input_file = input("Next Shriram file path: ").strip().strip('"')



# import pandas as pd
# import re, sys, os, json, time, urllib.request, urllib.error
# from datetime import datetime

# # =============================================================================
# #  SHRIRAM GENERAL INSURANCE — PayinConfig Generator  v3
# #  Reads master data from Masters_-_dec_2025.xlsx  +  JSON reference files
# #  Run:  python SHRIRAM_PayinConfig_v3.py
# # =============================================================================

# # ── Logging ──────────────────────────────────────────────────────────────────
# C = {
#     "OK"   : "\033[32m", "WARN" : "\033[33m", "ERR"  : "\033[31m",
#     "API"  : "\033[36m", "CACHE": "\033[90m", "INFO" : "",
# }
# _log_file = None

# def log(level, msg):
#     ts   = datetime.now().strftime("%H:%M:%S")
#     line = f"[{ts}][{level}] {msg}"
#     col  = C.get(level, "")
#     print(f"{col}{line}\033[0m" if col else line)
#     if _log_file:
#         _log_file.write(line + "\n")
#         _log_file.flush()

# def bar(cur, tot, w=48):
#     pct = cur / tot if tot else 0
#     b   = "\u2588" * int(w * pct) + "\u2591" * (w - int(w * pct))
#     print(f"\r  [{b}] {cur}/{tot} ({pct*100:.1f}%)", end="", flush=True)
#     if cur == tot: print()

# # ── Tiny helpers ──────────────────────────────────────────────────────────────
# def nt(v):
#     """Normalise: UPPER + collapse non-alnum to single space."""
#     return re.sub(r'\s+', ' ', re.sub(r'[^A-Z0-9]+', ' ', str(v).upper())).strip()

# def si(v, d=-1):
#     try:   return int(float(str(v)))
#     except: return d

# def sf(v, d=0.0):
#     if isinstance(v, str): v = v.strip().replace('%','')
#     try:   return float(v)
#     except: return d

# def load_dotenv(p):
#     if not os.path.exists(p): return
#     for line in open(p, encoding="utf-8"):
#         s = line.strip()
#         if not s or s.startswith('#') or '=' not in s: continue
#         k, v = s.split('=', 1)
#         k = k.strip(); v = v.strip().strip('"').strip("'")
#         if k and k not in os.environ: os.environ[k] = v

# # =============================================================================
# #  STARTUP
# # =============================================================================
# print("\n" + "="*75)
# print("  SHRIRAM GENERAL INSURANCE — PayinConfig Generator  v3")
# print("="*75)

# MASTERS = input("Path to Masters_-_dec_2025.xlsx  : ").strip().strip('"')
# JSON_DIR= input("Path to JSON files folder        : ").strip().strip('"')
# OUT_DIR = input("Output folder path               : ").strip().strip('"')
# os.makedirs(OUT_DIR, exist_ok=True)

# _log_path = os.path.join(OUT_DIR, f"shriram_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
# try:
#     _log_file = open(_log_path, "w", encoding="utf-8")
#     log("OK", f"Log → {_log_path}")
# except Exception as e:
#     print(f"[WARN] Cannot open log file: {e}")

# load_dotenv(os.path.join(JSON_DIR, ".env"))
# load_dotenv(".env")

# # =============================================================================
# #  LOAD MASTERS  (Excel sheets take priority; JSON files fill any gaps)
# # =============================================================================
# log("INFO", f"Loading master data from Excel: {MASTERS}")
# t0  = time.time()
# xl  = pd.ExcelFile(MASTERS)

# def rsh(name):
#     df = pd.read_excel(xl, sheet_name=name)
#     df.columns = [str(c).strip() for c in df.columns]
#     return df

# df_company = rsh("Insurance Company Master")   # company_id, company_code, company_name
# df_segment = rsh("Segment Master")            # id, segment
# df_subprod = rsh("SubProduct Master")         # sub_product_id, sub_product_name (+product_name)
# df_vtype   = rsh("Vehicle type")              # id, vehicle_type, no_of_wheels, sub_product_name
# df_fuel_xl = rsh("fuel master")               # id, fuel_type
# df_tw_mmv  = rsh("tw_pc_mmv master")          # MakeID/vehicle_make/ModelID/vehicle_model/VariantID/VariantDisplayName/CC/Fuel/SeatingCapacity/Is_Geared_Vehicle
# df_cv_mmv  = rsh("cv_mmv master")             # company_code/MakeID/vehicle_make/ModelID/vehicle_model/VariantID/VariantName/Fuel/SeatingCapacity

# log("OK", f"Excel masters loaded in {time.time()-t0:.1f}s")

# # ── Company dict ──────────────────────────────────────────────────────────────
# company_by_id = {si(r['company_id']): r for _, r in df_company.iterrows()}

# # ── Segment  (id → name, name → id) ──────────────────────────────────────────
# seg_name_to_id = {str(r['segment']).strip(): si(r['id']) for _, r in df_segment.iterrows()}
# # Policy-type string → canonical segment name
# PT_SEG = {
#     'COMPREHENSIVE': 'Comprehensive', 'COMP': 'Comprehensive',
#     'TP': 'TP Only',  'TP ONLY': 'TP Only', 'THIRD PARTY': 'TP Only',
#     'SAOD': 'SAOD',   'OD': 'SAOD',         'OWN DAMAGE': 'SAOD',
#     'TW NEW': 'Comprehensive',   # new business TW → Comprehensive
# }

# # ── Subproduct  (motor only, name → id) ──────────────────────────────────────
# df_sp_motor = df_subprod[df_subprod.get('product_name', pd.Series(dtype=str)).str.upper() == 'MOTOR'] \
#               if 'product_name' in df_subprod.columns else df_subprod
# subprod_id  = {str(r['sub_product_name']).strip(): si(r['sub_product_id'])
#                for _, r in df_sp_motor.iterrows()}

# # ── Vehicle type  (name → id, sub_product → default id) ──────────────────────
# vt_name_id = {str(r['vehicle_type']).strip(): si(r['id']) for _, r in df_vtype.iterrows()}
# vt_sub_default = {}   # sub_product_name → first vehicle_type_id for that sub
# for _, r in df_vtype.iterrows():
#     sp = str(r['sub_product_name']).strip()
#     if sp not in vt_sub_default:
#         vt_sub_default[sp] = si(r['id'])

# # ── Fuel  (name → id) ─────────────────────────────────────────────────────────
# fuel_id = {str(r['fuel_type']).strip().upper(): si(r['id']) for _, r in df_fuel_xl.iterrows()}
# # Normalise varied fuel strings from MMV sheets → canonical names
# _FUEL_NORM = {
#     'P':'PETROL','PETROL':'PETROL',
#     'D':'DIESEL','DIESEL':'DIESEL',
#     'ELECTRIC':'ELECTRIC','E':'ELECTRIC','B':'ELECTRIC',
#     'CNG':'CNG-LPG','LPG':'CNG-LPG','C':'CNG-LPG','CNG-LPG':'CNG-LPG',
#     'CNG/PETROL':'CNG-LPG','HYBRID':'PETROL','PETROL/ELECTRIC':'ELECTRIC',
# }
# def fuel_lookup(raw):
#     k = _FUEL_NORM.get(str(raw).strip().upper(), '')
#     return fuel_id.get(k, -1), k

# # =============================================================================
# #  MMV INDEX BUILD
# # =============================================================================
# log("INFO", "Building MMV indexes …")
# t1 = time.time()

# # ── TW / Private Car MMV  (global — no company scope) ─────────────────────────
# tw_make  = {}   # make_norm → (make_id, make_name)
# tw_model = {}   # (make_norm, model_norm) → (model_id, model_name)
# tw_var   = {}   # (make_norm, model_norm, var_norm) → (var_id, var_name, fuel_name, seating, cc, geared)
# tw_makes_sorted = []          # sorted list for substring scanning
# tw_models_of_make = {}        # make_norm → set of (model_norm, model_name, model_id)

# for _, m in df_tw_mmv.iterrows():
#     mkn = nt(m.get('vehicle_make',''));  mdn = nt(m.get('vehicle_model',''))
#     vrn = nt(m.get('VariantDisplayName',''))
#     mid = si(m.get('MakeID',-1)); mdid= si(m.get('ModelID',-1)); vid= si(m.get('VariantID',-1))
#     fid_v, fn = fuel_lookup(m.get('Fuel',''))
#     seat= si(m.get('SeatingCapacity',-1)); cc = si(m.get('CC',-1))
#     gear= 1 if str(m.get('Is_Geared_Vehicle','')).upper() in ('TRUE','1','YES') else 0

#     if mkn and mid != -1:
#         tw_make.setdefault(mkn, (mid, str(m.get('vehicle_make','')).strip()))
#     if mkn and mdn and mdid != -1:
#         tw_model.setdefault((mkn, mdn), (mdid, str(m.get('vehicle_model','')).strip()))
#         tw_models_of_make.setdefault(mkn, set()).add((mdn, str(m.get('vehicle_model','')).strip(), mdid))
#     if mkn and mdn and vrn and vid != -1:
#         tw_var.setdefault((mkn, mdn, vrn),
#                           (vid, str(m.get('VariantDisplayName','')).strip(), fn, seat, cc, gear))

# tw_makes_sorted = sorted([(k, tw_make[k][1], tw_make[k][0]) for k in tw_make],
#                          key=lambda x: len(x[0]), reverse=True)

# # ── CV MMV  (scoped by company_code) ──────────────────────────────────────────
# # IMPORTANT: Many makes appear with 2 different MakeIDs (duplicate batches in the master).
# # Strategy: pre-compute the DOMINANT MakeID per (company_code, vehicle_make) = the one
# # with the most variant rows. This is always the authoritative/most-complete batch.
# log("INFO", "Pre-computing dominant MakeID per company+make (resolves duplicate ID issue)...")
# _cv_make_row_count = {}   # (code_nt, make_nt, make_id) → row count
# for _, m in df_cv_mmv.iterrows():
#     code = nt(str(m.get('company_code','')))
#     mkn  = nt(str(m.get('vehicle_make','')))
#     mid  = si(m.get('MakeID', -1))
#     if code and mkn and mid != -1:
#         key = (code, mkn, mid)
#         _cv_make_row_count[key] = _cv_make_row_count.get(key, 0) + 1

# # For each (code, make_norm) pick the MakeID with the highest row count
# _dominant_make_id = {}   # (code, make_norm) → dominant make_id
# _all_make_counts  = {}   # (code, make_norm) → {make_id: count}
# for (code, mkn, mid), cnt in _cv_make_row_count.items():
#     _all_make_counts.setdefault((code, mkn), {})[mid] = cnt
# for (code, mkn), id_counts in _all_make_counts.items():
#     _dominant_make_id[(code, mkn)] = max(id_counts, key=id_counts.get)

# # Similarly pre-compute dominant ModelID per (code, make_norm, model_norm)
# _cv_model_row_count = {}
# for _, m in df_cv_mmv.iterrows():
#     code = nt(str(m.get('company_code','')))
#     mkn  = nt(str(m.get('vehicle_make','')))
#     mdn  = nt(str(m.get('vehicle_model','')))
#     mid  = si(m.get('MakeID', -1))
#     mdid = si(m.get('ModelID', -1))
#     if code and mkn and mdn and mdid != -1:
#         key = (code, mkn, mdn, mdid)
#         _cv_model_row_count[key] = _cv_model_row_count.get(key, 0) + 1

# _dominant_model_id = {}
# _all_model_counts  = {}
# for (code, mkn, mdn, mdid), cnt in _cv_model_row_count.items():
#     _all_model_counts.setdefault((code, mkn, mdn), {})[mdid] = cnt
# for (code, mkn, mdn), id_counts in _all_model_counts.items():
#     _dominant_model_id[(code, mkn, mdn)] = max(id_counts, key=id_counts.get)

# log("OK", f"Dominant IDs computed: {len(_dominant_make_id)} make keys, {len(_dominant_model_id)} model keys")

# cv_make  = {}   # (code, make_norm) → (make_id, make_name)
# cv_model = {}   # (code, make_norm, model_norm) → (model_id, model_name)
# cv_var   = {}   # (code, make_norm, model_norm, var_norm) → (var_id, var_name, fuel_name, seating)
# cv_makes_of_co  = {}   # code → sorted list of (make_norm, make_name, make_id)
# cv_models_of_mk = {}   # (code, make_norm) → set of (model_norm, model_name, model_id)
# cv_models_of_co = {}   # code → {model_norm: [(model_id, model_name, make_id, make_name)]}

# for _, m in df_cv_mmv.iterrows():
#     code= nt(m.get('company_code',''))
#     mkn = nt(m.get('vehicle_make',''));  mdn = nt(m.get('vehicle_model',''))
#     vrn = nt(m.get('VariantName',''))
#     mid = si(m.get('MakeID',-1)); mdid= si(m.get('ModelID',-1)); vid= si(m.get('VariantID',-1))
#     fid_v, fn = fuel_lookup(m.get('Fuel',''))
#     seat= si(m.get('SeatingCapacity',-1))
#     mk_display  = str(m.get('vehicle_make','')).strip()
#     md_display  = str(m.get('vehicle_model','')).strip()

#     if code and mkn and mid != -1:
#         dom_mid = _dominant_make_id.get((code, mkn), mid)
#         # Only index this row's make if mid == dominant, OR entry not yet set
#         if mid == dom_mid or (code, mkn) not in cv_make:
#             cv_make[(code, mkn)] = (dom_mid, mk_display)
#         cv_makes_of_co.setdefault(code, set()).add((mkn, mk_display, dom_mid))

#     if code and mkn and mdn and mdid != -1:
#         dom_mid  = _dominant_make_id.get((code, mkn), mid)
#         dom_mdid = _dominant_model_id.get((code, mkn, mdn), mdid)
#         if mdid == dom_mdid or (code, mkn, mdn) not in cv_model:
#             cv_model[(code, mkn, mdn)] = (dom_mdid, md_display)
#         cv_models_of_mk.setdefault((code, mkn), set()).add((mdn, md_display, dom_mdid))
#         cv_models_of_co.setdefault(code, {}).setdefault(mdn, [])
#         # Only add if not already present with same model_id
#         existing_ids = {x[0] for x in cv_models_of_co[code][mdn]}
#         if dom_mdid not in existing_ids:
#             cv_models_of_co[code][mdn].append((dom_mdid, md_display, dom_mid, mk_display))

#     if code and mkn and mdn and vrn and vid != -1:
#         cv_var.setdefault((code, mkn, mdn, vrn),
#                           (vid, str(m.get('VariantName','')).strip(), fn, seat))

# for k in list(cv_makes_of_co.keys()):
#     cv_makes_of_co[k] = sorted(list(cv_makes_of_co[k]), key=lambda x: len(x[0]), reverse=True)

# log("OK", (f"MMV indexes built in {time.time()-t1:.1f}s  |  "
#            f"TW/PC makes={len(tw_make)} models={len(tw_model)} variants={len(tw_var)}  |  "
#            f"CV make-rows={len(cv_make)} model-rows={len(cv_model)} variant-rows={len(cv_var)}"))

# # =============================================================================
# #  COMPANY SELECTION
# # =============================================================================
# print("\n" + "="*75)
# print("AVAILABLE COMPANIES")
# print("="*75)
# for cid, row in sorted(company_by_id.items()):
#     print(f"  {cid:3d}  |  {str(row['company_code']):20s}  |  {row['company_name']}")
# print("="*75)

# while True:
#     try:
#         CID      = int(input("\nEnter company_id: ").strip())
#         CROW     = company_by_id[CID]
#         CCODE    = str(CROW['company_code']).strip()
#         CCODE_NT = nt(CCODE)
#         log("OK", f"Company → {CROW['company_name']}  (id={CID}  code={CCODE})")
#         break
#     except (ValueError, KeyError):
#         log("ERR", "Invalid company_id, try again.")

# OUT_FILE = os.path.join(OUT_DIR, f"{CCODE}-Payin-Config.xlsx")

# _api_key = os.getenv("OPENAI_API_KEY","").strip()
# _model   = os.getenv("OPENAI_MODEL","gpt-4.1-mini")
# if _api_key: log("OK",   f"OpenAI key found — model: {_model}")
# else:        log("WARN", "No OPENAI_API_KEY — heuristic-only remark parsing")

# # =============================================================================
# #  PURE-PARSING HELPERS
# # =============================================================================

# def norm_pt(v):
#     """Policy type raw string → COMP / TP / SAOD."""
#     s = str(v).strip().upper()
#     if s in ('COMP','COMPREHENSIVE'):        return 'COMP'
#     if s in ('TP','TP ONLY','THIRD PARTY'):  return 'TP'
#     if s in ('SAOD','OD','OWN DAMAGE'):      return 'SAOD'
#     if s == 'TW NEW':                        return 'COMP'
#     return 'COMP'

# def vehicle_info(seg_text):
#     """SEGMENT column → (sub_product_name, sub_product_id, default_vehicle_type_id)."""
#     s = str(seg_text).strip().upper()
#     # Two Wheeler
#     if re.search(r'\bTW\b', s) or re.search(r'\b2W\b', s):
#         sp = 'Two Wheeler'
#         return sp, subprod_id.get(sp, -1), vt_name_id.get('TW Bike', vt_sub_default.get(sp, -1))
#     # Private Car
#     if re.search(r'\bPVT\s*CAR\b', s) or re.search(r'\bPRIVATE\s*CAR\b', s) or re.search(r'\bPC\b', s):
#         sp = 'Private Car'
#         return sp, subprod_id.get(sp, -1), vt_name_id.get('Private Car', vt_sub_default.get(sp, -1))
#     # Passenger Vehicle
#     if re.search(r'\bPCV\b', s) or re.search(r'\bPASSENGER\b', s):
#         sp = 'Passenger Vehicle'
#         return sp, subprod_id.get(sp, -1), vt_name_id.get('Auto rikshaw', vt_sub_default.get(sp, -1))
#     # Miscellaneous — Tractor / Harvester
#     if 'TRACTOR' in s or 'HARVESTER' in s:
#         sp = 'Miscellaneous Vehicle'
#         return sp, subprod_id.get(sp, -1), vt_name_id.get('Agriculture Tractor', vt_sub_default.get(sp, -1))
#     # Miscellaneous — MISD / MISC
#     if 'MISD' in s or 'MISC' in s:
#         sp = 'Miscellaneous Vehicle'
#         return sp, subprod_id.get(sp, -1), vt_name_id.get('Non Tractor', vt_sub_default.get(sp, -1))
#     # Goods Vehicle — GCV / GVW / CV
#     sp = 'Goods Vehicle'
#     return sp, subprod_id.get(sp, -1), vt_name_id.get('Truck', vt_sub_default.get(sp, -1))

# def parse_age(s):
#     s = str(s).strip()
#     if s.lower() in ('', 'nan', 'none', 'new', 'n/a'): return 0, 700
#     su = s.upper()
#     m = re.search(r'UPTO\s+(\d+)\s*(YEAR|YR|MONTH|MTH|MO)', su)
#     if m: n, u = int(m.group(1)), m.group(2); return 0, n*12 if u.startswith('Y') else n
#     m = re.search(r'(\d+)\s*TO\s*(\d+)\s*(YEAR|YR)?', su)
#     if m: return int(m.group(1))*12, int(m.group(2))*12
#     m = re.match(r'>\s*(\d+)\s*[-]\s*(\d+)(\+?)\s*[Yy]', s)
#     if m: return int(m.group(1))*12+1, (700 if m.group(3)=='+' else int(m.group(2))*12)
#     m = re.match(r'>\s*(\d+)\+?\s*[Yy]', s)
#     if m: return int(m.group(1))*12+1, 700
#     m = re.match(r'^(\d+)$', s.strip())
#     if m: return 0, int(m.group(1))*12
#     return 0, 700

# def parse_cc(s):
#     if not s or str(s).strip().lower() in ('', 'nan', 'none'): return 0, 99999, -1
#     s = str(s).strip().upper().replace('CC','').strip()
#     for pat, fn in [
#         (r'^<\s*(\d+)$',                 lambda m: (0, int(m.group(1))-1, 1)),
#         (r'^>\s*(\d+)$',                 lambda m: (int(m.group(1))+1, 99999, 1)),
#         (r'^>\s*(\d+)\s*[-]\s*(\d+)$',  lambda m: (int(m.group(1))+1, int(m.group(2)), 1)),
#         (r'^>=\s*(\d+)\s*[-]\s*(\d+)$', lambda m: (int(m.group(1)), int(m.group(2)), 1)),
#         (r'^(\d+)\s*[-]\s*(\d+)$',      lambda m: (int(m.group(1)), int(m.group(2)), 1)),
#         (r'^(\d+)$',                     lambda m: (int(m.group(1)), int(m.group(1)), 1)),
#     ]:
#         mt = re.match(pat, s)
#         if mt: return fn(mt)
#     return 0, 99999, -1

# def parse_idv(text):
#     su = str(text).upper()
#     m = re.search(r'IDV\s+(?:UPTO|UP\s*TO|<=?)\s*([\d.]+)\s*(?:LAC|LAKH|L\b)', su)
#     if m: return 1, 0.0, float(m.group(1))
#     m = re.search(r'IDV\s+([\d.]+)\s*[-]\s*([\d.]+)\s*(?:LAC|LAKH)', su)
#     if m: return 1, float(m.group(1)), float(m.group(2))
#     return -1, 0.0, 0.0

# def parse_weight(text):
#     su = str(text).upper()
#     m = re.search(r'([\d.]+)\s*KG\b', su)
#     if m: return 1, 0, int(float(m.group(1)))
#     m = re.search(r'UPTO\s+([\d.]+)\s*T(?:ON|ONNE)?', su)
#     if m: return 1, 0, int(float(m.group(1))*1000)
#     m = re.search(r'>\s*([\d.]+)\s*[-]\s*([\d.]+)\s*T(?:ON|ONNE)?', su)
#     if m: return 1, int(float(m.group(1))*1000)+1, int(float(m.group(2))*1000)
#     m = re.search(r'\b([\d.]+)\s*T(?:ON|ONNE)?\b', su)
#     if m: return 1, 0, int(float(m.group(1))*1000)
#     m = re.search(r'GVW\s+([\d.]+)', su)
#     if m: v = float(m.group(1)); return 1, 0, int(v*1000) if v < 200 else int(v)
#     return -1, 0, 99999

# def parse_seating(text):
#     su = str(text).upper()
#     m = re.search(r'(?:SC|SEATING\s*CAPACITY|SEATING)\s*UPTO\s*(\d+)', su)
#     if m: return 1, 1, int(m.group(1))
#     m = re.search(r'(?:SC|SEATING\s*CAPACITY|SEATING)\s*(\d+)\s*[-]\s*(\d+)', su)
#     if m: return 1, int(m.group(1)), int(m.group(2))
#     m = re.search(r'(?:SC|SEATING\s*CAPACITY|SEATING)\s*(\d+)', su)
#     if m: v = int(m.group(1)); return 1, v, v
#     m = re.search(r'(\d+)\s*SEATER', su)
#     if m: v = int(m.group(1)); return 1, v, v
#     return -1, -1, -1

# def parse_fuel_text(text):
#     su = str(text).upper()
#     if 'ELECTRIC' in su or ' EV ' in su: return fuel_id.get('ELECTRIC',-1), 'ELECTRIC'
#     if 'CNG' in su or 'LPG' in su:       return fuel_id.get('CNG-LPG',-1),  'CNG-LPG'
#     if 'DIESEL' in su:                   return fuel_id.get('DIESEL',-1),    'DIESEL'
#     if 'PETROL' in su:                   return fuel_id.get('PETROL',-1),    'PETROL'
#     return -1, ''

# def ncb_flag(text):
#     su = str(text).upper()
#     if any(x in su for x in ('WITHOUT NCB','W/O NCB','NON NCB','NON-NCB','ZERO NCB')): return 0
#     if 'NCB' in su: return 1
#     return -1

# def irda_flag(text):
#     su = str(text).upper()
#     return 1 if any(x in su for x in ('IRDA TP','IRDA RATE','IRDA')) else -1

# def cpa_flag(text):
#     su = str(text).upper()
#     if 'CPA' in su and any(x in su for x in ('INCLUD','WITH CPA')): return 1
#     return -1

# def keep_included(text):
#     """Return only the 'included' portion — strip excluded/except/rejected chunks."""
#     s = str(text).strip()
#     if not s or s.lower() in ('nan', 'none'): return ''
#     kept = []
#     for chunk in re.split(r';', s):
#         up = chunk.upper()
#         if any(t in up for t in ('DECLIN','REJECT')): continue
#         for tok in (' BUT ',' EXCEPT ',' EXCLUDE ',' OTHER THAN ',' NOT CONSIDER'):
#             idx = up.find(tok)
#             if idx != -1: chunk = chunk[:idx]; break
#         chunk = chunk.strip()
#         if chunk: kept.append(chunk)
#     return ' '.join(x for x in kept if x)

# def is_pure_exclusion(text):
#     """True if the remark has ONLY exclusion language and no inclusion items."""
#     s = str(text).strip().upper()
#     if not s or s in ('NAN','NONE'): return False
#     EXCL = ('EXCEPT','EXCLUDE','OTHER THAN','REJECT','DECLIN','NOT CONSIDER',
#             'HR 68','EXCLUDED')
#     INCL = ('ONLY','INCLUD','HONDA','BAJAJ','TATA','MARUTI','HYUNDAI','KIA',
#             'MAHINDRA','IDV','NCB','ZONE','BRANCH')
#     if any(t in s for t in INCL): return False
#     return any(t in s for t in EXCL)

# # =============================================================================
# #  MMV RESOLUTION
# # =============================================================================

# def _score_make_match(query_norm, candidate_norm):
#     """
#     Score how well candidate_norm matches query_norm.
#     Higher = better match. Returns -1 if no match at all.

#     Scoring priority (descending):
#       100 : exact match  (query == candidate)
#        80 : all query words found in candidate AND candidate starts with query
#        60 : all query words are a subset of candidate words  (query words ⊆ candidate words)
#        40 : query is a substring of candidate  (e.g. "TATA" in "TATA MOTORS")
#        20 : candidate is a substring of query  (e.g. "TATA" matches query "TATA MAKE ONLY")
#        -1 : no match

#     Tiebreaker: prefer candidate with FEWER extra words (closer length to query).
#     """
#     if not query_norm or not candidate_norm:
#         return -1, 0
#     qw = set(query_norm.split())
#     cw = set(candidate_norm.split())
#     extra = len(cw) - len(qw)   # extra words in candidate beyond query (smaller = better)

#     if query_norm == candidate_norm:
#         return 100, -extra
#     if qw.issubset(cw) and candidate_norm.startswith(query_norm):
#         return 80, -extra
#     if qw.issubset(cw):
#         return 60, -extra
#     if query_norm in candidate_norm:
#         return 40, -extra
#     if candidate_norm in query_norm:
#         return 20, -extra
#     return -1, 0

# def _best_make_match(mkn, candidates):
#     """
#     candidates = list of (norm_name, display_name, make_id)
#     Returns (make_id, display_name, matched_norm) or (-1, '', '')
#     """
#     best_score = -1; best_tie = 0; best = None
#     for cn, cname, cid in candidates:
#         score, tie = _score_make_match(mkn, cn)
#         if score > best_score or (score == best_score and tie > best_tie):
#             best_score, best_tie, best = score, tie, (cid, cname, cn)
#     if best_score >= 20:   # accept any real match
#         return best
#     return -1, '', ''


# def resolve_tw(raw_make, raw_model, raw_variant, remark):
#     mkn = nt(raw_make) if raw_make else ''
#     mdn = nt(raw_model) if raw_model else ''
#     vrn = nt(raw_variant) if raw_variant else ''
#     mk_id=-1; mk_name=''; md_id=-1; md_name=''; vr_id=-1; vr_name=''
#     fuel=''; seat=-1; cc=-1; gear=-1
#     inc = nt(keep_included(remark))

#     # 1. Make — direct hit first, then scored scan
#     if mkn:
#         if mkn in tw_make:
#             mk_id, mk_name = tw_make[mkn]
#         else:
#             mk_id, mk_name, mkn = _best_make_match(mkn, tw_makes_sorted)
#         if mk_id == -1:  # last resort: scan included remark text
#             for cn, cname, cid in tw_makes_sorted:
#                 if cn and re.search(rf'\b{re.escape(cn)}\b', inc):
#                     mk_id, mk_name, mkn = cid, cname, cn; break

#     # 2. Model (only if mentioned)
#     if mdn and mk_id != -1:
#         if (mkn, mdn) in tw_model:
#             md_id, md_name = tw_model[(mkn, mdn)]
#         else:
#             model_cands = [(mn, mname, mid) for mn,mname,mid
#                            in tw_models_of_make.get(mkn, set())]
#             md_id, md_name, mdn = _best_make_match(mdn, model_cands)

#     # 3. Variant (only if mentioned)
#     if vrn and mk_id != -1 and md_id != -1:
#         key = (mkn, mdn, vrn)
#         if key in tw_var:
#             vr_id, vr_name, fuel, seat, cc, gear = tw_var[key]
#         else:
#             for (cmn,cmdn,cvn), (vid,vname,vf,vs,vc,vg) in tw_var.items():
#                 if cmn==mkn and cmdn==mdn and (vrn in cvn or cvn in vrn):
#                     vr_id,vr_name,fuel,seat,cc,gear = vid,vname,vf,vs,vc,vg; break

#     # Infer fuel/seating from first matching variant when no variant given
#     if mk_id != -1 and md_id != -1 and not vrn:
#         for (cmn,cmdn,_),(vid,vname,vf,vs,vc,vg) in tw_var.items():
#             if cmn==mkn and cmdn==mdn and vf:
#                 fuel,seat,cc,gear = vf,vs,vc,vg; break

#     return mk_id, mk_name, md_id, md_name, vr_id, vr_name, fuel, seat, cc, gear


# def resolve_cv(raw_make, raw_model, raw_variant, remark):
#     code = CCODE_NT
#     mkn  = nt(raw_make)    if raw_make    else ''
#     mdn  = nt(raw_model)   if raw_model   else ''
#     vrn  = nt(raw_variant) if raw_variant else ''
#     mk_id=-1; mk_name=''; md_id=-1; md_name=''; vr_id=-1; vr_name=''
#     fuel=''; seat=-1
#     inc = nt(keep_included(remark))

#     # 1. Make — direct hit first, then scored scan
#     if mkn:
#         if (code, mkn) in cv_make:
#             mk_id, mk_name = cv_make[(code, mkn)]
#         else:
#             mk_id, mk_name, mkn = _best_make_match(mkn, cv_makes_of_co.get(code, []))
#         if mk_id == -1:  # last resort: scan included remark text
#             for cn, cname, cid in cv_makes_of_co.get(code, []):
#                 if cn and re.search(rf'\b{re.escape(cn)}\b', inc):
#                     mk_id, mk_name, mkn = cid, cname, cn; break

#     # 2. Model — direct hit first, then scored scan
#     if mdn:
#         if mk_id != -1 and (code, mkn, mdn) in cv_model:
#             md_id, md_name = cv_model[(code, mkn, mdn)]
#         elif mk_id != -1:
#             model_cands = [(mn, mname, mid) for mn,mname,mid
#                            in cv_models_of_mk.get((code, mkn), set())]
#             md_id, md_name, mdn = _best_make_match(mdn, model_cands)
#         if md_id == -1:  # model without known make
#             hits = cv_models_of_co.get(code, {}).get(mdn, [])
#             if hits: md_id, md_name, mk_id, mk_name = hits[0]

#     # 3. Variant
#     if vrn and mk_id != -1 and md_id != -1:
#         key = (code, mkn, mdn, vrn)
#         if key in cv_var:
#             vr_id, vr_name, fuel, seat = cv_var[key]
#         else:
#             for (ck,cmn,cmdn,cvn),(vid,vname,vf,vs) in cv_var.items():
#                 if ck==code and cmn==mkn and cmdn==mdn and (vrn in cvn or cvn in vrn):
#                     vr_id,vr_name,fuel,seat = vid,vname,vf,vs; break

#     if mk_id != -1 and md_id != -1 and not vrn:
#         for (ck,cmn,cmdn,_),(vid,vname,vf,vs) in cv_var.items():
#             if ck==code and cmn==mkn and cmdn==mdn and vf:
#                 fuel, seat = vf, vs; break

#     return mk_id, mk_name, md_id, md_name, vr_id, vr_name, fuel, seat

# # =============================================================================
# #  OPENAI REMARK PARSER
# # =============================================================================

# def heuristic(remark, seg, pt):
#     su = str(remark).upper()
#     fid,fn  = parse_fuel_text(remark)
#     iw,fw,tw_= parse_weight(remark)
#     iss,fs,ts= parse_seating(remark)
#     ii,fi,ti = parse_idv(remark)
#     return {
#         'vehicle_make':'','vehicle_model':'','vehicle_variant':'',
#         'is_with_ncb':ncb_flag(su), 'is_irda_tp':irda_flag(su),
#         'is_cpa_included':cpa_flag(su),
#         'fuel_type':fn,
#         'seating_cap':-1,'from_seating':fs,'to_seating':ts,
#         'is_weight':iw,'from_weight_kg':fw,'to_weight_kg':tw_,
#         'idv_cap':ii,'from_idv':fi,'to_idv':ti,
#         'is_cc':-1,'from_cc':0,'to_cc':99999,
#     }

# _nc=0; _nh=0; _ne=0; _ms=0

# def parse_remark(remark, co_name, seg, pt, row_n=0):
#     global _nc, _ne, _ms
#     ak = os.getenv("OPENAI_API_KEY","").strip()
#     if not ak: return heuristic(remark, seg, pt)

#     included = keep_included(remark)
#     short    = (remark[:80]+'…') if len(remark)>80 else remark
#     log("API", f"Row {row_n:>4} | call #{_nc+1} | seg={seg!r:20} | {short!r}")

#     prompt = f"""You are an expert Indian motor insurance data extractor.
# Analyse the remark carefully. Return ONLY a valid JSON object — no markdown, no preamble — with EXACTLY these keys:

#   vehicle_make    : INCLUDED makes comma-separated (e.g. "HONDA,HYUNDAI,KIA"). Empty if none.
#   vehicle_model   : INCLUDED models comma-separated. Empty if none.
#   vehicle_variant : Specific variant name if mentioned, else "".
#   is_with_ncb     : 1 if NCB cases INCLUDED. 0 if WITHOUT/NON/ZERO NCB. -1 if not mentioned.
#   is_irda_tp      : 1 if IRDA TP rate mentioned, -1 otherwise.
#   is_cpa_included : 1 if CPA included/mentioned, -1 otherwise.
#   fuel_type       : DIESEL | PETROL | ELECTRIC | CNG-LPG | "" (empty if not mentioned).
#   seating_cap     : exact integer if single seating value given, -1 otherwise.
#   from_seating    : lower seating bound (int), -1 if N/A.
#   to_seating      : upper seating bound (int), -1 if N/A.
#   is_weight       : 1 if GVW/weight/tonnage mentioned, -1 otherwise.
#   from_weight_kg  : lower weight KG int, 0 if N/A.
#   to_weight_kg    : upper weight KG int, 99999 if N/A.
#   idv_cap         : 1 if IDV cap mentioned, -1 otherwise.
#   from_idv        : lower IDV in Lacs (float), 0 if N/A.
#   to_idv          : upper IDV in Lacs (float), 0 if N/A.
#   is_cc           : 1 if engine CC/capacity mentioned, -1 otherwise.
#   from_cc         : lower CC int, 0 if N/A.
#   to_cc           : upper CC int, 99999 if N/A.

# RULES — READ CAREFULLY:
# 1. ONLY list makes/models that are INCLUDED. IGNORE anything after EXCEPT / BUT / EXCLUDE /
#    OTHER THAN / REJECT / DECLINE / ONLY EXCEPT.
# 2. If remark has ONLY exclusion text (e.g. "Except TATA" or "HR 68 EXCLUDED"), set
#    vehicle_make="" and vehicle_model="".
# 3. SC = Seating Capacity. "SC 7" → seating_cap=7. "SC upto 7" → from_seating=1, to_seating=7.
# 4. Tons → KG: 1 ton = 1000 KG. "7.5T" → to_weight_kg=7500.
# 5. "IDV upto 10 lacs" → idv_cap=1, from_idv=0, to_idv=10.
# 6. "Upto 1500 CC" → is_cc=1, from_cc=0, to_cc=1500.
# 7. TRACTOR segment → vehicle_variant="Agriculture Tractor".

# company: {co_name}
# segment: {seg}
# policy_type: {pt}
# remark_original: {remark}
# remark_included_only: {included}
# """
#     body = {"model": _model,
#             "messages": [
#                 {"role":"system","content":"Return only valid JSON, no markdown, no explanation."},
#                 {"role":"user","content":prompt}],
#             "response_format":{"type":"json_object"}, "temperature":0}
#     req = urllib.request.Request(
#         "https://api.openai.com/v1/chat/completions",
#         data=json.dumps(body).encode(),
#         headers={"Content-Type":"application/json","Authorization":f"Bearer {ak}"},
#         method="POST")
#     ts = time.time()
#     try:
#         with urllib.request.urlopen(req, timeout=30) as r:
#             p = json.loads(json.loads(r.read())["choices"][0]["message"]["content"])
#         ms = int((time.time()-ts)*1000); _nc+=1; _ms+=ms
#         def _s(k,d=""): return str(p.get(k,d)).strip()
#         def _i(k,d=-1):
#             try: return int(p.get(k,d))
#             except: return d
#         def _f(k,d=0.0):
#             try: return float(p.get(k,d))
#             except: return d
#         result = {
#             'vehicle_make':_s('vehicle_make'),'vehicle_model':_s('vehicle_model'),
#             'vehicle_variant':_s('vehicle_variant'),
#             'is_with_ncb':_i('is_with_ncb'),'is_irda_tp':_i('is_irda_tp'),
#             'is_cpa_included':_i('is_cpa_included'),
#             'fuel_type':_s('fuel_type').upper(),
#             'seating_cap':_i('seating_cap'),'from_seating':_i('from_seating'),
#             'to_seating':_i('to_seating'),
#             'is_weight':_i('is_weight'),
#             'from_weight_kg':_i('from_weight_kg',0),'to_weight_kg':_i('to_weight_kg',99999),
#             'idv_cap':_i('idv_cap'),'from_idv':_f('from_idv'),'to_idv':_f('to_idv'),
#             'is_cc':_i('is_cc'),'from_cc':_i('from_cc',0),'to_cc':_i('to_cc',99999),
#         }
#         log("OK", (f"Row {row_n:>4} | {ms}ms | "
#                    f"make={result['vehicle_make']!r:15} model={result['vehicle_model']!r:12} "
#                    f"ncb={result['is_with_ncb']} irda={result['is_irda_tp']} "
#                    f"cc={result['is_cc']} fuel={result['fuel_type']!r}"))
#         return result
#     except urllib.error.HTTPError as e:
#         _ne+=1; log("ERR", f"Row {row_n:>4} | HTTP {e.code} → heuristic")
#     except Exception as e:
#         _ne+=1; log("ERR", f"Row {row_n:>4} | {e} → heuristic")
#     return heuristic(remark, seg, pt)

# # =============================================================================
# #  OUTPUT COLUMN ORDER
# # =============================================================================
# COLS = [
#     'id','company_id','company_code','segment_id','segment',
#     'subproduct_id','sub_product_name','lob_id','lob_name',
#     'business_type_id','business_type','is_highend_lob',
#     'rto_group_id','rto_group_name',
#     'payin_od_rate','payin_tp_rate','payout_od_rate','payout_tp_rate',
#     'extra_tp_rate','eff_from_date','eff_to_date',
#     'fuel_type_id','fuel_type',
#     'is_on_net','is_one_year_pay_on_newbusiness','is_cpa_included',
#     'is_geared_vehicle','is_cc_considered','from_cc','to_cc',
#     'is_premium_considered','from_premium','to_premium',
#     'is_mmv_considered','make_id','vehicle_make','model_id','vehicle_model',
#     'variant_id','vehicle_variant',
#     'is_seating_cap_consider','from_seating_cap','to_seating_cap',
#     'is_no_of_wheel_consider','from_no_of_wheel','to_no_of_wheel',
#     'vehicle_type_id','ppi_in','ppi_out',
#     'is_irda_tp_included','is_longterm_renewal_pay',
#     'is_weightage_considered','from_weightage_kg','to_weightage_kg',
#     'is_nil_dep_considered','is_organization_type',
#     'from_age_month','to_age_month',
#     'is_with_ncb','is_idv_cap_consider','from_idv','to_idv',
#     'is_breakin_consider','is_active',
# ]

# def make_row(
#     sid,sname,spid,spname,rto_id,rto_name,
#     pod,ptp,pod2,ptp2,
#     fid,fname,is_on_net,cpa,geared,
#     iscc,fcc,tcc,
#     iswt,fwt,twt,
#     vtid,
#     is_mmv,mkid,mkname,mdid,mdname,vrid,vrname,
#     issc,fsc,tsc,
#     nw,fnw,tnw,
#     fage,tage,
#     ncb,irda,
#     idvcap,fidv,tidv,
# ):
#     return {
#         'id':0, 'company_id':CID, 'company_code':CCODE,
#         'segment_id':sid, 'segment':sname,
#         'subproduct_id':spid, 'sub_product_name':spname,
#         'lob_id':-1, 'lob_name':'',
#         'business_type_id':-1, 'business_type':'Not Considered', 'is_highend_lob':False,
#         'rto_group_id':rto_id, 'rto_group_name':rto_name,
#         'payin_od_rate':pod, 'payin_tp_rate':ptp,
#         'payout_od_rate':pod2, 'payout_tp_rate':ptp2,
#         'extra_tp_rate':0, 'eff_from_date':'2026-01-01', 'eff_to_date':'2026-01-16',
#         'fuel_type_id':fid, 'fuel_type':fname,
#         'is_on_net':is_on_net,
#         'is_one_year_pay_on_newbusiness':-1, 'is_cpa_included':cpa,
#         'is_geared_vehicle':geared,
#         'is_cc_considered':iscc, 'from_cc':fcc, 'to_cc':tcc,
#         'is_premium_considered':-1, 'from_premium':-1, 'to_premium':-1,
#         'is_mmv_considered':is_mmv,
#         'make_id':mkid, 'vehicle_make':mkname,
#         'model_id':mdid, 'vehicle_model':mdname,
#         'variant_id':vrid, 'vehicle_variant':vrname,
#         'is_seating_cap_consider':issc, 'from_seating_cap':fsc, 'to_seating_cap':tsc,
#         'is_no_of_wheel_consider':nw, 'from_no_of_wheel':fnw, 'to_no_of_wheel':tnw,
#         'vehicle_type_id':vtid, 'ppi_in':0, 'ppi_out':0,
#         'is_irda_tp_included':irda, 'is_longterm_renewal_pay':-1,
#         'is_weightage_considered':iswt, 'from_weightage_kg':fwt, 'to_weightage_kg':twt,
#         'is_nil_dep_considered':-1, 'is_organization_type':-1,
#         'from_age_month':fage, 'to_age_month':tage,
#         'is_with_ncb':ncb,
#         'is_idv_cap_consider':idvcap, 'from_idv':fidv, 'to_idv':tidv,
#         'is_breakin_consider':-1, 'is_active':True,
#     }

# # =============================================================================
# #  PROCESS ONE INPUT FILE
# # =============================================================================

# def process(input_file):
#     global _nh
#     log("INFO", f"Reading: {input_file}")
#     tr = time.time()
#     df = pd.read_excel(input_file)
#     df.columns = [c.strip() for c in df.columns]
#     log("OK", f"Loaded {len(df)} rows in {time.time()-tr:.1f}s | cols: {list(df.columns)}")

#     def col(*names):
#         for n in names:
#             if n in df.columns: return n
#         return None

#     c_seg  = col('SEGMENT','Segment')
#     c_pt   = col('POLICY TYPE','Policy Type','POLICYTYPE')
#     c_loc  = col('LOCATION','Location')
#     c_pay  = col('PAYIN','Payin')
#     c_pout = col('PAYOUT','Payout','Calculated Payout')
#     c_rem  = col('REMARK','Remark','REMARKS','CALCULATION EXPLANATION')
#     c_age  = col('AGE','Age','AGE BAND')
#     c_cc   = col('CC BAND','CC Band','CC')
#     c_tw   = col('TW TYPE','TW Type')
#     c_co   = col('COMPANY NAME','Company Name','COMPANY')

#     log("INFO", f"col-map → seg={c_seg!r} pt={c_pt!r} loc={c_loc!r} "
#         f"pay={c_pay!r} pout={c_pout!r} rem={c_rem!r} age={c_age!r}")

#     # estimate API calls
#     uniq_rem = set()
#     if c_rem:
#         for v in df[c_rem].fillna('').astype(str): uniq_rem.add(v.strip())
#     log("INFO", f"Unique remarks: {len(uniq_rem)} → "
#         + (f"~{len(uniq_rem)} API calls" if _api_key else "heuristic only"))

#     out_rows = []; cache = {}; total = len(df); tp = time.time()
#     print(f"\n  {'='*58}\n  Processing {total} rows …\n  {'='*58}\n")

#     for idx, (_, row) in enumerate(df.iterrows(), 1):
#         bar(idx, total)

#         if idx == 1 or idx % 25 == 0 or idx == total:
#             el = time.time()-tp; rate = idx/el if el else 0; eta = (total-idx)/rate if rate else 0
#             avg = (_ms/_nc) if _nc else 0
#             log("INFO", (f"Row {idx:>4}/{total} | {el:.0f}s | ETA {eta:.0f}s | "
#                          f"rate={rate:.1f}/s | api={_nc} avg={avg:.0f}ms | "
#                          f"cache={_nh} err={_ne} | out={len(out_rows)}"))

#         def g(c, d=''):
#             if c is None: return d
#             v = row.get(c, d)
#             return v if v is not None and str(v).strip() not in ('nan','None','NaN') else d

#         # ── Policy type & segment ─────────────────────────────────────────────
#         pt_raw = str(g(c_pt, 'COMP')).strip()
#         pt     = norm_pt(pt_raw)
#         seg_nm = PT_SEG.get(pt, 'Comprehensive')
#         seg_id = seg_name_to_id.get(seg_nm, 1)

#         # ── Rates ─────────────────────────────────────────────────────────────
#         payin  = sf(g(c_pay, 0))
#         payout = sf(g(c_pout, 0))
#         if   pt == 'TP':   pod, ptp, pod2, ptp2 = 0, payin, 0, payout
#         elif pt == 'SAOD': pod, ptp, pod2, ptp2 = payin, 0, payout, 0
#         else:              pod = ptp = payin; pod2 = ptp2 = payout

#         # ── Vehicle / subproduct ──────────────────────────────────────────────
#         seg_text = str(g(c_seg, '')).strip()
#         spname, spid, vtid = vehicle_info(seg_text)
#         is_cv   = spname in ('Goods Vehicle','Passenger Vehicle','Miscellaneous Vehicle')

#         # ── Location (rto_group_id always 0 per corrections.txt) ─────────────
#         rto_name = str(g(c_loc, '')).strip()
#         rto_id   = 0

#         # ── Other columns ─────────────────────────────────────────────────────
#         remark   = str(g(c_rem, '')).strip()
#         co_name  = str(g(c_co, CCODE)).strip()
#         fage, tage = parse_age(str(g(c_age, '')))
#         fcc0, tcc0, iscc0 = parse_cc(str(g(c_cc, '')))

#         # TW geared / vehicle-type override
#         tw_raw = str(g(c_tw, '')).strip().lower()
#         geared = -1
#         if spname == 'Two Wheeler' and tw_raw:
#             if 'scooter'  in tw_raw: vtid = vt_name_id.get('TW Scooter', vtid); geared = 0
#             elif 'bike'   in tw_raw: vtid = vt_name_id.get('TW Bike',    vtid); geared = 1
#             elif 'electric' in tw_raw: vtid = vt_name_id.get('TW Electric Bike', vtid); geared = -1

#         # ── is_on_net ─────────────────────────────────────────────────────────
#         is_on_net = True if pt == 'COMP' else False

#         # ── OpenAI / cache ────────────────────────────────────────────────────
#         ck = (remark, co_name, seg_text, pt)
#         if ck in cache:
#             _nh += 1
#             log("CACHE", f"Row {idx:>4} | HIT #{_nh} | {remark[:50]!r}")
#             meta = cache[ck]
#         else:
#             meta  = parse_remark(remark, co_name, seg_text, pt, idx)
#             cache[ck] = meta

#         # ── Pull meta fields ──────────────────────────────────────────────────
#         ncb    = meta['is_with_ncb']
#         irda   = meta['is_irda_tp']
#         cpa    = meta['is_cpa_included']
#         raw_mk = meta['vehicle_make']
#         raw_md = meta['vehicle_model']
#         raw_vr = meta['vehicle_variant']

#         # Fuel
#         m_fuel = meta.get('fuel_type', '')
#         if m_fuel and m_fuel in fuel_id: fid_v, fn_v = fuel_id[m_fuel], m_fuel
#         else: fid_v, fn_v = parse_fuel_text(remark)

#         # Seating
#         msc  = meta.get('seating_cap',-1)
#         mfsc = meta.get('from_seating',-1)
#         mtsc = meta.get('to_seating',-1)
#         if msc != -1:                        issc_v,fsc_v,tsc_v = 1,msc,msc
#         elif mfsc != -1 or mtsc != -1:       issc_v=1; fsc_v=mfsc if mfsc!=-1 else 1; tsc_v=mtsc if mtsc!=-1 else 99
#         else:                                issc_v,fsc_v,tsc_v = parse_seating(remark)

#         # Weight
#         m_wt  = meta.get('is_weight',-1)
#         if m_wt == 1: iswt_v,fwt_v,twt_v = 1,meta['from_weight_kg'],meta['to_weight_kg']
#         else:         iswt_v,fwt_v,twt_v = parse_weight(remark)

#         # IDV
#         m_idv = meta.get('idv_cap',-1)
#         if m_idv == 1: idv_v,fidv_v,tidv_v = 1,meta['from_idv'],meta['to_idv']
#         else:          idv_v,fidv_v,tidv_v = parse_idv(remark)

#         # CC  (column-level overrides OpenAI if column exists)
#         iscc_v, fcc_v, tcc_v = iscc0, fcc0, tcc0
#         if iscc_v == -1 and meta.get('is_cc',-1) == 1:
#             iscc_v, fcc_v, tcc_v = 1, meta['from_cc'], meta['to_cc']

#         # Pure exclusion → clear MMV
#         if is_pure_exclusion(remark):
#             raw_mk = raw_md = raw_vr = ''
#             log("INFO", f"Row {idx:>4} | pure-exclusion remark → MMV cleared")

#         # ── Expand multiple makes ─────────────────────────────────────────────
#         make_list = [m.strip() for m in re.split(r'[,&]+', raw_mk) if m.strip()] or ['']
#         if len(make_list) > 1:
#             log("INFO", f"Row {idx:>4} | expanding {len(make_list)} makes: {make_list}")

#         for one_make in make_list:
#             if is_cv:
#                 mkid,mkname,mdid,mdname,vrid,vrname,i_fuel,i_seat = \
#                     resolve_cv(one_make, raw_md, raw_vr, remark)
#                 i_cc=-1; i_gear=-1
#             else:
#                 mkid,mkname,mdid,mdname,vrid,vrname,i_fuel,i_seat,i_cc,i_gear = \
#                     resolve_tw(one_make, raw_md, raw_vr, remark)

#             if one_make and mkid == -1:
#                 log("WARN", f"Row {idx:>4} | make '{one_make}' NOT in MMV for {CCODE}")

#             # Tractor → vehicle_variant override
#             if spname == 'Miscellaneous Vehicle' and 'TRACTOR' in seg_text.upper():
#                 if not vrname: vrname = 'Agriculture Tractor'

#             is_mmv = 1 if (mkid!=-1 or mdid!=-1 or vrid!=-1 or
#                            one_make or raw_md or raw_vr) else -1

#             # Fuel fallback from MMV variant data
#             fin_fid, fin_fn = fid_v, fn_v
#             if fin_fid == -1 and i_fuel and i_fuel in fuel_id:
#                 fin_fid = fuel_id[i_fuel]; fin_fn = i_fuel

#             # Seating fallback from MMV
#             fin_issc, fin_fsc, fin_tsc = issc_v, fsc_v, tsc_v
#             if fin_issc == -1 and i_seat > 0:
#                 fin_issc=1; fin_fsc=i_seat; fin_tsc=i_seat

#             # Geared fallback from MMV variant
#             fin_gear = geared
#             if spname == 'Two Wheeler' and fin_gear == -1 and i_gear != -1:
#                 fin_gear = i_gear

#             out_rows.append(make_row(
#                 seg_id, seg_nm, spid, spname, rto_id, rto_name,
#                 pod, ptp, pod2, ptp2,
#                 fin_fid, fin_fn, is_on_net, cpa, fin_gear,
#                 iscc_v, fcc_v, tcc_v,
#                 iswt_v, fwt_v, twt_v,
#                 vtid,
#                 is_mmv,
#                 mkid, mkname if mkname else one_make,
#                 mdid, mdname if mdname else raw_md,
#                 vrid, vrname if vrname else raw_vr,
#                 fin_issc, fin_fsc, fin_tsc,
#                 -1, -1, -1,
#                 fage, tage,
#                 ncb, irda,
#                 idv_v, fidv_v, tidv_v,
#             ))

#     el = time.time()-tp; avg = (_ms/_nc) if _nc else 0
#     print()
#     log("OK","="*60)
#     log("OK",f"DONE  input={total}  output={len(out_rows)}  time={el:.1f}s ({el/60:.1f}min)")
#     log("OK",f"  API calls={_nc}  avg={avg:.0f}ms  cache={_nh}  errors={_ne}")
#     log("OK","="*60)
#     return pd.DataFrame(out_rows)

# # =============================================================================
# #  MAIN LOOP  — process one or more files, all appended to same output
# # =============================================================================
# input_file = input("\nEnter Shriram input Excel file path: ").strip().strip('"')

# while True:
#     try:
#         df_out = process(input_file)
#         df_out = df_out[[c for c in COLS if c in df_out.columns]]

#         if os.path.exists(OUT_FILE):
#             log("INFO", f"Appending to existing: {OUT_FILE}")
#             df_out = pd.concat([pd.read_excel(OUT_FILE), df_out], ignore_index=True)

#         df_out.to_excel(OUT_FILE, index=False)
#         log("OK", f"Saved → {OUT_FILE}  ({len(df_out)} rows)")
#         log("OK", f"Log   → {_log_path}")

#     except Exception as e:
#         import traceback
#         log("ERR", f"FATAL: {e}")
#         traceback.print_exc()

#     print("\n" + "="*75)
#     print("  1  Process another Shriram file (appends to same output)")
#     print("  2  Exit")
#     ch = input("Choice: ").strip()
#     if ch == "2":
#         log("OK","Goodbye!")
#         if _log_file: _log_file.close()
#         break
#     input_file = input("Next Shriram file path: ").strip().strip('"')
"""
PayinConfig Generator — Universal Motor Insurance Pay-in Config Tool
Supports Shriram and any other company in the master list.

Features:
  - Loads master data from Masters_-_dec_2025.xlsx (Excel sheets)
  - OpenAI-powered remark parsing with heuristic fallback
  - Make/Model/Variant resolution for TW/PC (global) and CV (per company)
  - Dominant-ID strategy for CV MMV to avoid duplicate MakeID issues
  - Full column schema as per spec
  - Progress bar + colored logging + log file
  - Append mode: multiple input files -> single output
  - Remark caching to avoid duplicate API calls

Usage:
  python PayinConfig.py
  Set OPENAI_API_KEY in environment or in a .env file in the JSON folder.
"""

# import pandas as pd
# import re, sys, os, json, time, urllib.request, urllib.error
# from datetime import datetime

# # =============================================================================
# #  LOGGING
# # =============================================================================
# COLORS = {
#     "OK"   : "\033[32m",
#     "WARN" : "\033[33m",
#     "ERR"  : "\033[31m",
#     "API"  : "\033[36m",
#     "CACHE": "\033[90m",
#     "INFO" : "",
# }
# _log_file = None

# def log(level, msg):
#     ts   = datetime.now().strftime("%H:%M:%S")
#     line = f"[{ts}][{level}] {msg}"
#     col  = COLORS.get(level, "")
#     print(f"{col}{line}\033[0m" if col else line)
#     if _log_file:
#         _log_file.write(line + "\n")
#         _log_file.flush()

# def progress_bar(cur, tot, width=50):
#     pct = cur / tot if tot else 0
#     filled = int(width * pct)
#     bar = "█" * filled + "░" * (width - filled)
#     print(f"\r  [{bar}] {cur}/{tot} ({pct*100:.1f}%)", end="", flush=True)
#     if cur == tot:
#         print()

# # =============================================================================
# #  UTILITIES
# # =============================================================================

# def nt(v):
#     """Normalize: UPPER-CASE and collapse non-alphanumeric to single space."""
#     return re.sub(r'\s+', ' ', re.sub(r'[^A-Z0-9]+', ' ', str(v).upper())).strip()

# def si(v, default=-1):
#     try:   return int(float(str(v)))
#     except: return default

# def sf(v, default=0.0):
#     if isinstance(v, str):
#         v = v.strip().replace('%', '')
#     try:   return float(v)
#     except: return default

# def load_dotenv(path):
#     if not os.path.exists(path):
#         return
#     with open(path, encoding="utf-8") as f:
#         for line in f:
#             s = line.strip()
#             if not s or s.startswith('#') or '=' not in s:
#                 continue
#             k, v = s.split('=', 1)
#             k = k.strip()
#             v = v.strip().strip('"').strip("'")
#             if k and k not in os.environ:
#                 os.environ[k] = v

# # =============================================================================
# #  STARTUP
# # =============================================================================
# print("\n" + "="*75)
# print("  PayinConfig Generator  — Motor Insurance Pay-in Config Tool")
# print("="*75)

# MASTERS  = input("Path to Masters_-_dec_2025.xlsx : ").strip().strip('"')
# JSON_DIR = input("Path to JSON files folder       : ").strip().strip('"')
# OUT_DIR  = input("Output folder path              : ").strip().strip('"')
# os.makedirs(OUT_DIR, exist_ok=True)

# _log_path = os.path.join(OUT_DIR, f"payin_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
# try:
#     _log_file = open(_log_path, "w", encoding="utf-8")
#     log("OK", f"Log → {_log_path}")
# except Exception as e:
#     print(f"[WARN] Cannot open log file: {e}")

# load_dotenv(os.path.join(JSON_DIR, ".env"))
# load_dotenv(".env")

# # =============================================================================
# #  LOAD MASTER DATA FROM EXCEL
# # =============================================================================
# log("INFO", f"Loading master data: {MASTERS}")
# t0 = time.time()
# xl = pd.ExcelFile(MASTERS)

# def read_sheet(name):
#     df = pd.read_excel(xl, sheet_name=name)
#     df.columns = [str(c).strip() for c in df.columns]
#     return df

# df_company = read_sheet("Insurance Company Master")
# df_segment = read_sheet("Segment Master")
# df_subprod = read_sheet("SubProduct Master")
# df_vtype   = read_sheet("Vehicle type")
# df_fuel    = read_sheet("fuel master")
# df_tw_mmv  = read_sheet("tw_pc_mmv master")
# df_cv_mmv  = read_sheet("cv_mmv master")

# log("OK", f"Masters loaded in {time.time()-t0:.1f}s")

# # ── Company lookup ──────────────────────────────────────────────────────────
# company_by_id = {si(r['company_id']): r for _, r in df_company.iterrows()}

# # ── Segment lookup ──────────────────────────────────────────────────────────
# seg_name_to_id = {str(r['segment']).strip(): si(r['id'])
#                   for _, r in df_segment.iterrows()}

# POLICY_TO_SEG = {
#     'COMP': 'Comprehensive',  'COMPREHENSIVE': 'Comprehensive',
#     'TP':   'TP Only',        'TP ONLY':       'TP Only',   'THIRD PARTY': 'TP Only',
#     'SAOD': 'SAOD',           'OD':            'SAOD',      'OWN DAMAGE':  'SAOD',
#     'TW NEW': 'Comprehensive',
# }

# # ── Subproduct lookup (motor only) ─────────────────────────────────────────
# if 'product_name' in df_subprod.columns:
#     df_sp = df_subprod[df_subprod['product_name'].str.upper() == 'MOTOR']
# else:
#     df_sp = df_subprod
# subprod_id = {str(r['sub_product_name']).strip(): si(r['sub_product_id'])
#               for _, r in df_sp.iterrows()}

# # ── Vehicle type lookup ─────────────────────────────────────────────────────
# vt_name_to_id  = {str(r['vehicle_type']).strip(): si(r['id'])
#                   for _, r in df_vtype.iterrows()}
# vt_sub_default = {}  # sub_product_name → first vehicle_type_id
# for _, r in df_vtype.iterrows():
#     sp = str(r['sub_product_name']).strip()
#     if sp not in vt_sub_default:
#         vt_sub_default[sp] = si(r['id'])

# # ── Fuel lookup ─────────────────────────────────────────────────────────────
# fuel_id = {str(r['fuel_type']).strip().upper(): si(r['id'])
#            for _, r in df_fuel.iterrows()}

# _FUEL_ALIASES = {
#     'PETROL': 'PETROL',   'P': 'PETROL',
#     'DIESEL': 'DIESEL',   'D': 'DIESEL',
#     'ELECTRIC': 'ELECTRIC', 'E': 'ELECTRIC', 'EV': 'ELECTRIC', 'B': 'ELECTRIC',
#     'CNG-LPG': 'CNG-LPG', 'CNG': 'CNG-LPG', 'LPG': 'CNG-LPG', 'C': 'CNG-LPG',
#     'CNG/PETROL': 'CNG-LPG', 'HYBRID': 'PETROL', 'PETROL/ELECTRIC': 'ELECTRIC',
# }

# def fuel_lookup(raw):
#     """Return (fuel_type_id, fuel_type_name) from raw string."""
#     k = _FUEL_ALIASES.get(str(raw).strip().upper(), '')
#     return fuel_id.get(k, -1), k

# # =============================================================================
# #  MMV INDEX BUILD
# # =============================================================================
# log("INFO", "Building MMV indexes …")
# t1 = time.time()

# # ── TW / Private Car MMV (global, no company scope) ──────────────────────────
# tw_make           = {}  # norm_make → (make_id, display_make)
# tw_model          = {}  # (norm_make, norm_model) → (model_id, display_model)
# tw_var            = {}  # (norm_make, norm_model, norm_var) → (var_id, display_var, fuel, seat, cc, geared)
# tw_makes_sorted   = []  # [(norm, display, id)] sorted by len(norm) desc
# tw_models_of_make = {}  # norm_make → set of (norm_model, display_model, model_id)

# for _, m in df_tw_mmv.iterrows():
#     mkn = nt(m.get('vehicle_make', ''))
#     mdn = nt(m.get('vehicle_model', ''))
#     vrn = nt(m.get('VariantDisplayName', ''))
#     mid  = si(m.get('MakeID', -1))
#     mdid = si(m.get('ModelID', -1))
#     vid  = si(m.get('VariantID', -1))
#     fid_v, fn = fuel_lookup(m.get('Fuel', ''))
#     seat = si(m.get('SeatingCapacity', -1))
#     cc   = si(m.get('CC', -1))
#     gear = 1 if str(m.get('Is_Geared_Vehicle', '')).upper() in ('TRUE', '1', 'YES') else 0

#     if mkn and mid != -1:
#         tw_make.setdefault(mkn, (mid, str(m.get('vehicle_make', '')).strip()))
#     if mkn and mdn and mdid != -1:
#         tw_model.setdefault((mkn, mdn), (mdid, str(m.get('vehicle_model', '')).strip()))
#         tw_models_of_make.setdefault(mkn, set()).add(
#             (mdn, str(m.get('vehicle_model', '')).strip(), mdid))
#     if mkn and mdn and vrn and vid != -1:
#         tw_var.setdefault((mkn, mdn, vrn),
#                           (vid, str(m.get('VariantDisplayName', '')).strip(), fn, seat, cc, gear))

# tw_makes_sorted = sorted(
#     [(k, tw_make[k][1], tw_make[k][0]) for k in tw_make],
#     key=lambda x: len(x[0]), reverse=True)

# # ── CV MMV (per company code) ─────────────────────────────────────────────────
# # Pre-compute dominant MakeID/ModelID to handle duplicate batches in master
# _cv_make_cnt = {}   # (code, make_norm, make_id) → row count
# _cv_model_cnt = {}  # (code, make_norm, model_norm, model_id) → row count

# for _, m in df_cv_mmv.iterrows():
#     code = nt(str(m.get('company_code', '')))
#     mkn  = nt(str(m.get('vehicle_make', '')))
#     mdn  = nt(str(m.get('vehicle_model', '')))
#     mid  = si(m.get('MakeID', -1))
#     mdid = si(m.get('ModelID', -1))
#     if code and mkn and mid != -1:
#         _cv_make_cnt[(code, mkn, mid)] = _cv_make_cnt.get((code, mkn, mid), 0) + 1
#     if code and mkn and mdn and mdid != -1:
#         _cv_model_cnt[(code, mkn, mdn, mdid)] = _cv_model_cnt.get((code, mkn, mdn, mdid), 0) + 1

# # For each (code, make_norm) choose MakeID with most rows
# _dom_make_id  = {}
# _make_agg = {}
# for (code, mkn, mid), cnt in _cv_make_cnt.items():
#     _make_agg.setdefault((code, mkn), {})[mid] = cnt
# for (code, mkn), id_cnt in _make_agg.items():
#     _dom_make_id[(code, mkn)] = max(id_cnt, key=id_cnt.get)

# # For each (code, make_norm, model_norm) choose ModelID with most rows
# _dom_model_id = {}
# _model_agg = {}
# for (code, mkn, mdn, mdid), cnt in _cv_model_cnt.items():
#     _model_agg.setdefault((code, mkn, mdn), {})[mdid] = cnt
# for (code, mkn, mdn), id_cnt in _model_agg.items():
#     _dom_model_id[(code, mkn, mdn)] = max(id_cnt, key=id_cnt.get)

# cv_make          = {}  # (code, norm_make) → (make_id, display_make)
# cv_model         = {}  # (code, norm_make, norm_model) → (model_id, display_model)
# cv_var           = {}  # (code, norm_make, norm_model, norm_var) → (var_id, display_var, fuel, seat)
# cv_makes_of_co   = {}  # code → sorted [(norm, display, id)]
# cv_models_of_mk  = {}  # (code, norm_make) → set of (norm_model, display_model, model_id)
# cv_models_of_co  = {}  # code → {norm_model: [(model_id, display_model, make_id, display_make)]}

# for _, m in df_cv_mmv.iterrows():
#     code = nt(m.get('company_code', ''))
#     mkn  = nt(m.get('vehicle_make', ''))
#     mdn  = nt(m.get('vehicle_model', ''))
#     vrn  = nt(m.get('VariantName', ''))
#     mid  = si(m.get('MakeID', -1))
#     mdid = si(m.get('ModelID', -1))
#     vid  = si(m.get('VariantID', -1))
#     fid_v, fn = fuel_lookup(m.get('Fuel', ''))
#     seat = si(m.get('SeatingCapacity', -1))
#     mk_disp = str(m.get('vehicle_make', '')).strip()
#     md_disp = str(m.get('vehicle_model', '')).strip()

#     if code and mkn and mid != -1:
#         dom_mid = _dom_make_id.get((code, mkn), mid)
#         if mid == dom_mid or (code, mkn) not in cv_make:
#             cv_make[(code, mkn)] = (dom_mid, mk_disp)
#         cv_makes_of_co.setdefault(code, set()).add((mkn, mk_disp, dom_mid))

#     if code and mkn and mdn and mdid != -1:
#         dom_mid  = _dom_make_id.get((code, mkn), mid)
#         dom_mdid = _dom_model_id.get((code, mkn, mdn), mdid)
#         if mdid == dom_mdid or (code, mkn, mdn) not in cv_model:
#             cv_model[(code, mkn, mdn)] = (dom_mdid, md_disp)
#         cv_models_of_mk.setdefault((code, mkn), set()).add((mdn, md_disp, dom_mdid))
#         cv_models_of_co.setdefault(code, {}).setdefault(mdn, [])
#         existing = {x[0] for x in cv_models_of_co[code][mdn]}
#         if dom_mdid not in existing:
#             cv_models_of_co[code][mdn].append((dom_mdid, md_disp, dom_mid, mk_disp))

#     if code and mkn and mdn and vrn and vid != -1:
#         cv_var.setdefault((code, mkn, mdn, vrn),
#                           (vid, str(m.get('VariantName', '')).strip(), fn, seat))

# for k in cv_makes_of_co:
#     cv_makes_of_co[k] = sorted(list(cv_makes_of_co[k]),
#                                 key=lambda x: len(x[0]), reverse=True)

# log("OK", (f"MMV indexes built in {time.time()-t1:.1f}s | "
#            f"TW makes={len(tw_make)} models={len(tw_model)} variants={len(tw_var)} | "
#            f"CV makes={len(cv_make)} models={len(cv_model)} variants={len(cv_var)}"))

# # =============================================================================
# #  COMPANY SELECTION
# # =============================================================================
# print("\n" + "="*75)
# print("AVAILABLE COMPANIES")
# print("="*75)
# for cid, row in sorted(company_by_id.items()):
#     print(f"  {cid:3d}  |  {str(row['company_code']):20s}  |  {row['company_name']}")
# print("="*75)

# while True:
#     try:
#         CID      = int(input("\nEnter company_id: ").strip())
#         CROW     = company_by_id[CID]
#         CCODE    = str(CROW['company_code']).strip()
#         CCODE_NT = nt(CCODE)
#         log("OK", f"Company → {CROW['company_name']}  (id={CID}  code={CCODE})")
#         break
#     except (ValueError, KeyError):
#         log("ERR", "Invalid company_id — please try again.")

# OUT_FILE = os.path.join(OUT_DIR, f"{CCODE}-Payin-Config.xlsx")

# _api_key = os.getenv("OPENAI_API_KEY", "").strip()
# _ai_model = os.getenv("OPENAI_MODEL", "gpt-4.1-mini")
# if _api_key:
#     log("OK",   f"OpenAI key found — model: {_ai_model}")
# else:
#     log("WARN", "No OPENAI_API_KEY — heuristic remark parsing only")

# # =============================================================================
# #  PARSING HELPERS
# # =============================================================================

# def normalize_policy_type(raw):
#     s = str(raw).strip().upper()
#     if s in ('COMP', 'COMPREHENSIVE'):           return 'COMP'
#     if s in ('TP', 'TP ONLY', 'THIRD PARTY'):   return 'TP'
#     if s in ('SAOD', 'OD', 'OWN DAMAGE'):       return 'SAOD'
#     if s == 'TW NEW':                            return 'COMP'
#     return 'COMP'

# def vehicle_info_from_segment(seg_text):
#     """
#     Infer (sub_product_name, sub_product_id, vehicle_type_id) from SEGMENT column.
#     Handles TW, PC, PCV, GCV, CV, Misc, Tractor, etc.
#     """
#     s = str(seg_text).strip().upper()

#     if re.search(r'\bTW\b', s) or re.search(r'\b2W\b', s):
#         sp = 'Two Wheeler'
#         return sp, subprod_id.get(sp, -1), vt_name_to_id.get('TW Bike', vt_sub_default.get(sp, -1))

#     if re.search(r'\bPVT\s*CAR\b', s) or re.search(r'\bPRIVATE\s*CAR\b', s) or re.search(r'\bPC\b', s):
#         sp = 'Private Car'
#         return sp, subprod_id.get(sp, -1), vt_name_to_id.get('Private Car', vt_sub_default.get(sp, -1))

#     if re.search(r'\bPCV\b', s) or re.search(r'\bPASSENGER\b', s):
#         sp = 'Passenger Vehicle'
#         return sp, subprod_id.get(sp, -1), vt_name_to_id.get('Auto rikshaw', vt_sub_default.get(sp, -1))

#     if 'TRACTOR' in s or 'HARVESTER' in s:
#         sp = 'Miscellaneous Vehicle'
#         return sp, subprod_id.get(sp, -1), vt_name_to_id.get('Agriculture Tractor', vt_sub_default.get(sp, -1))

#     if 'MISD' in s or 'MISC' in s:
#         sp = 'Miscellaneous Vehicle'
#         return sp, subprod_id.get(sp, -1), vt_name_to_id.get('Non Tractor', vt_sub_default.get(sp, -1))

#     # Default: Goods Vehicle (GCV / GVW / CV)
#     sp = 'Goods Vehicle'
#     return sp, subprod_id.get(sp, -1), vt_name_to_id.get('Truck', vt_sub_default.get(sp, -1))

# def parse_age(s):
#     """Parse age field → (from_age_month, to_age_month)."""
#     s = str(s).strip()
#     if s.lower() in ('', 'nan', 'none', 'new', 'n/a'):
#         return 0, 700
#     su = s.upper()

#     m = re.search(r'UPTO\s+(\d+)\s*(YEAR|YR|MONTH|MTH|MO)', su)
#     if m:
#         n, u = int(m.group(1)), m.group(2)
#         return 0, n * 12 if u.startswith('Y') else n

#     m = re.search(r'(\d+)\s*TO\s*(\d+)\s*(YEAR|YR)?', su)
#     if m:
#         return int(m.group(1)) * 12, int(m.group(2)) * 12

#     m = re.match(r'>\s*(\d+)\s*[-]\s*(\d+)(\+?)\s*[Yy]', s)
#     if m:
#         lo = int(m.group(1)) * 12 + 1
#         hi = 700 if m.group(3) == '+' else int(m.group(2)) * 12
#         return lo, hi

#     m = re.match(r'>\s*(\d+)\+?\s*[Yy]', s)
#     if m:
#         return int(m.group(1)) * 12 + 1, 700

#     m = re.match(r'^(\d+)$', s.strip())
#     if m:
#         return 0, int(m.group(1)) * 12

#     return 0, 700

# def parse_cc_band(s):
#     """Parse CC band field → (from_cc, to_cc, is_cc_considered)."""
#     if not s or str(s).strip().lower() in ('', 'nan', 'none'):
#         return 0, 99999, -1
#     s = str(s).strip().upper().replace('CC', '').strip()
#     patterns = [
#         (r'^<\s*(\d+)$',                 lambda m: (0, int(m.group(1)) - 1, 1)),
#         (r'^>\s*(\d+)$',                 lambda m: (int(m.group(1)) + 1, 99999, 1)),
#         (r'^>\s*(\d+)\s*[-]\s*(\d+)$',  lambda m: (int(m.group(1)) + 1, int(m.group(2)), 1)),
#         (r'^>=\s*(\d+)\s*[-]\s*(\d+)$', lambda m: (int(m.group(1)), int(m.group(2)), 1)),
#         (r'^(\d+)\s*[-]\s*(\d+)$',      lambda m: (int(m.group(1)), int(m.group(2)), 1)),
#         (r'^(\d+)$',                     lambda m: (int(m.group(1)), int(m.group(1)), 1)),
#     ]
#     for pat, fn in patterns:
#         mt = re.match(pat, s)
#         if mt:
#             return fn(mt)
#     return 0, 99999, -1

# def parse_idv(text):
#     """Parse IDV cap from remark → (is_idv_cap, from_idv, to_idv)."""
#     su = str(text).upper()
#     m = re.search(r'IDV\s+(?:UPTO|UP\s*TO|<=?)\s*([\d.]+)\s*(?:LAC|LAKH|L\b)', su)
#     if m:
#         return 1, 0.0, float(m.group(1))
#     m = re.search(r'IDV\s+([\d.]+)\s*[-]\s*([\d.]+)\s*(?:LAC|LAKH)', su)
#     if m:
#         return 1, float(m.group(1)), float(m.group(2))
#     return -1, 0.0, 0.0

# def parse_weight(text):
#     """Parse weight/tonnage from remark → (is_weight, from_kg, to_kg)."""
#     su = str(text).upper()
#     m = re.search(r'([\d.]+)\s*KG\b', su)
#     if m:
#         return 1, 0, int(float(m.group(1)))
#     m = re.search(r'UPTO\s+([\d.]+)\s*T(?:ON|ONNE)?', su)
#     if m:
#         return 1, 0, int(float(m.group(1)) * 1000)
#     m = re.search(r'>\s*([\d.]+)\s*[-]\s*([\d.]+)\s*T(?:ON|ONNE)?', su)
#     if m:
#         return 1, int(float(m.group(1)) * 1000) + 1, int(float(m.group(2)) * 1000)
#     m = re.search(r'\b([\d.]+)\s*T(?:ON|ONNE)?\b', su)
#     if m:
#         return 1, 0, int(float(m.group(1)) * 1000)
#     m = re.search(r'GVW\s+([\d.]+)', su)
#     if m:
#         v = float(m.group(1))
#         return 1, 0, int(v * 1000) if v < 200 else int(v)
#     return -1, 0, 99999

# def parse_seating(text):
#     """Parse seating capacity from remark → (is_seating, from_sc, to_sc)."""
#     su = str(text).upper()
#     m = re.search(r'(?:SC|SEATING\s*CAPACITY|SEATING)\s*UPTO\s*(\d+)', su)
#     if m:
#         return 1, 1, int(m.group(1))
#     m = re.search(r'(?:SC|SEATING\s*CAPACITY|SEATING)\s*(\d+)\s*[-]\s*(\d+)', su)
#     if m:
#         return 1, int(m.group(1)), int(m.group(2))
#     m = re.search(r'(?:SC|SEATING\s*CAPACITY|SEATING)\s*\(?\s*(\d+)', su)
#     if m:
#         v = int(m.group(1))
#         return 1, v, v
#     m = re.search(r'(\d+)\s*SEATER', su)
#     if m:
#         v = int(m.group(1))
#         return 1, v, v
#     return -1, -1, -1

# def parse_fuel_from_text(text):
#     """Detect fuel type from free text → (fuel_type_id, fuel_type_name)."""
#     su = str(text).upper()
#     if 'ELECTRIC' in su or ' EV ' in su:
#         return fuel_id.get('ELECTRIC', -1), 'ELECTRIC'
#     if 'CNG' in su or 'LPG' in su:
#         return fuel_id.get('CNG-LPG', -1), 'CNG-LPG'
#     if 'DIESEL' in su:
#         return fuel_id.get('DIESEL', -1), 'DIESEL'
#     if 'PETROL' in su:
#         return fuel_id.get('PETROL', -1), 'PETROL'
#     return -1, ''

# def ncb_flag(text):
#     su = str(text).upper()
#     if any(x in su for x in ('WITHOUT NCB', 'W/O NCB', 'NON NCB', 'NON-NCB', 'ZERO NCB')):
#         return 0
#     if 'NCB' in su:
#         return 1
#     return -1

# def irda_flag(text):
#     su = str(text).upper()
#     return 1 if any(x in su for x in ('IRDA TP', 'IRDA RATE', 'IRDA')) else -1

# def cpa_flag(text):
#     su = str(text).upper()
#     if 'CPA' in su and any(x in su for x in ('INCLUD', 'WITH CPA')):
#         return 1
#     return -1

# def nil_dep_flag(text):
#     su = str(text).upper()
#     if 'NIL DEP' in su or 'NIL DEPRECIATION' in su or 'ZERO DEP' in su:
#         return 1
#     return -1

# def keep_included_only(text):
#     """Strip excluded/rejected clauses; return only the included portion."""
#     s = str(text).strip()
#     if not s or s.lower() in ('nan', 'none'):
#         return ''
#     kept = []
#     for chunk in re.split(r';', s):
#         up = chunk.upper()
#         if any(t in up for t in ('DECLIN', 'REJECT')):
#             continue
#         for tok in (' BUT ', ' EXCEPT ', ' EXCLUDE ', ' OTHER THAN ', ' NOT CONSIDER'):
#             idx = up.find(tok)
#             if idx != -1:
#                 chunk = chunk[:idx]
#                 break
#         chunk = chunk.strip()
#         if chunk:
#             kept.append(chunk)
#     return ' '.join(x for x in kept if x)

# def is_pure_exclusion(text):
#     """Return True if remark contains ONLY exclusion language."""
#     s = str(text).strip().upper()
#     if not s or s in ('NAN', 'NONE', ''):
#         return False
#     EXCL_TOKENS = ('EXCEPT', 'EXCLUDE', 'OTHER THAN', 'REJECT', 'DECLIN',
#                    'NOT CONSIDER', 'HR 68', 'EXCLUDED')
#     INCL_TOKENS = ('ONLY', 'INCLUD', 'HONDA', 'BAJAJ', 'TATA', 'MARUTI',
#                    'HYUNDAI', 'KIA', 'MAHINDRA', 'IDV', 'NCB', 'ZONE',
#                    'BRANCH', 'NCB CASES')
#     if any(t in s for t in INCL_TOKENS):
#         return False
#     return any(t in s for t in EXCL_TOKENS)

# # =============================================================================
# #  MMV RESOLUTION
# # =============================================================================

# def _score_match(query, candidate):
#     """
#     Score how well candidate matches query (both normalized).
#     Returns (score, tie_breaker). Higher = better.
#     100: exact  80: all words + starts with  60: words subset
#     40: query in candidate  20: candidate in query  -1: no match
#     """
#     if not query or not candidate:
#         return -1, 0
#     qw = set(query.split())
#     cw = set(candidate.split())
#     extra = len(cw) - len(qw)

#     if query == candidate:
#         return 100, -extra
#     if qw.issubset(cw) and candidate.startswith(query):
#         return 80, -extra
#     if qw.issubset(cw):
#         return 60, -extra
#     if query in candidate:
#         return 40, -extra
#     if candidate in query:
#         return 20, -extra
#     return -1, 0

# def _best_match(query_norm, candidates):
#     """
#     candidates: list of (norm, display, id)
#     Returns (id, display, matched_norm) or (-1, '', '')
#     """
#     best_score, best_tie, best = -1, 0, None
#     for cn, cname, cid in candidates:
#         score, tie = _score_match(query_norm, cn)
#         if score > best_score or (score == best_score and tie > best_tie):
#             best_score, best_tie, best = score, tie, (cid, cname, cn)
#     if best_score >= 20:
#         return best
#     return -1, '', ''


# def resolve_tw_mmv(raw_make, raw_model, raw_variant, remark):
#     """
#     Resolve TW/PC make/model/variant from MMV master.
#     Returns (make_id, make_name, model_id, model_name, var_id, var_name,
#              fuel_name, seat, cc, geared)
#     """
#     mkn = nt(raw_make) if raw_make else ''
#     mdn = nt(raw_model) if raw_model else ''
#     vrn = nt(raw_variant) if raw_variant else ''
#     mk_id = mk_name = -1, ''
#     md_id = md_name = -1, ''
#     vr_id = vr_name = -1, ''
#     fuel = ''; seat = cc = gear = -1
#     mk_id, mk_name = -1, ''
#     md_id, md_name = -1, ''
#     vr_id, vr_name = -1, ''
#     inc = nt(keep_included_only(remark))

#     # 1. Make
#     if mkn:
#         if mkn in tw_make:
#             mk_id, mk_name = tw_make[mkn]
#         else:
#             mk_id, mk_name, mkn = _best_match(mkn, tw_makes_sorted)
#         if mk_id == -1:
#             for cn, cname, cid in tw_makes_sorted:
#                 if cn and re.search(rf'\b{re.escape(cn)}\b', inc):
#                     mk_id, mk_name, mkn = cid, cname, cn
#                     break

#     # 2. Model
#     if mdn and mk_id != -1:
#         if (mkn, mdn) in tw_model:
#             md_id, md_name = tw_model[(mkn, mdn)]
#         else:
#             cands = [(mn, mname, mid) for mn, mname, mid in tw_models_of_make.get(mkn, set())]
#             md_id, md_name, mdn = _best_match(mdn, cands)

#     # 3. Variant
#     if vrn and mk_id != -1 and md_id != -1:
#         key = (mkn, mdn, vrn)
#         if key in tw_var:
#             vr_id, vr_name, fuel, seat, cc, gear = tw_var[key]
#         else:
#             for (cmn, cmdn, cvn), (vid, vname, vf, vs, vc, vg) in tw_var.items():
#                 if cmn == mkn and cmdn == mdn and (vrn in cvn or cvn in vrn):
#                     vr_id, vr_name, fuel, seat, cc, gear = vid, vname, vf, vs, vc, vg
#                     break

#     # Infer fuel/seat/gear from first variant when no variant specified
#     if mk_id != -1 and md_id != -1 and not vrn:
#         for (cmn, cmdn, _), (vid, vname, vf, vs, vc, vg) in tw_var.items():
#             if cmn == mkn and cmdn == mdn and vf:
#                 fuel, seat, cc, gear = vf, vs, vc, vg
#                 break

#     return mk_id, mk_name, md_id, md_name, vr_id, vr_name, fuel, seat, cc, gear


# def resolve_cv_mmv(raw_make, raw_model, raw_variant, remark):
#     """
#     Resolve CV make/model/variant from MMV master (scoped by CCODE_NT).
#     Returns (make_id, make_name, model_id, model_name, var_id, var_name, fuel, seat)
#     """
#     code = CCODE_NT
#     mkn  = nt(raw_make) if raw_make else ''
#     mdn  = nt(raw_model) if raw_model else ''
#     vrn  = nt(raw_variant) if raw_variant else ''
#     mk_id, mk_name = -1, ''
#     md_id, md_name = -1, ''
#     vr_id, vr_name = -1, ''
#     fuel = ''; seat = -1
#     inc  = nt(keep_included_only(remark))

#     # 1. Make
#     if mkn:
#         if (code, mkn) in cv_make:
#             mk_id, mk_name = cv_make[(code, mkn)]
#         else:
#             mk_id, mk_name, mkn = _best_match(mkn, cv_makes_of_co.get(code, []))
#         if mk_id == -1:
#             for cn, cname, cid in cv_makes_of_co.get(code, []):
#                 if cn and re.search(rf'\b{re.escape(cn)}\b', inc):
#                     mk_id, mk_name, mkn = cid, cname, cn
#                     break

#     # 2. Model
#     if mdn:
#         if mk_id != -1 and (code, mkn, mdn) in cv_model:
#             md_id, md_name = cv_model[(code, mkn, mdn)]
#         elif mk_id != -1:
#             cands = [(mn, mname, mid) for mn, mname, mid in cv_models_of_mk.get((code, mkn), set())]
#             md_id, md_name, mdn = _best_match(mdn, cands)
#         if md_id == -1:
#             hits = cv_models_of_co.get(code, {}).get(mdn, [])
#             if hits:
#                 md_id, md_name, mk_id, mk_name = hits[0]

#     # 3. Variant
#     if vrn and mk_id != -1 and md_id != -1:
#         key = (code, mkn, mdn, vrn)
#         if key in cv_var:
#             vr_id, vr_name, fuel, seat = cv_var[key]
#         else:
#             for (ck, cmn, cmdn, cvn), (vid, vname, vf, vs) in cv_var.items():
#                 if ck == code and cmn == mkn and cmdn == mdn and (vrn in cvn or cvn in vrn):
#                     vr_id, vr_name, fuel, seat = vid, vname, vf, vs
#                     break

#     if mk_id != -1 and md_id != -1 and not vrn:
#         for (ck, cmn, cmdn, _), (vid, vname, vf, vs) in cv_var.items():
#             if ck == code and cmn == mkn and cmdn == mdn and vf:
#                 fuel, seat = vf, vs
#                 break

#     return mk_id, mk_name, md_id, md_name, vr_id, vr_name, fuel, seat

# # =============================================================================
# #  OPENAI REMARK PARSER
# # =============================================================================

# def heuristic_parse(remark, seg, pt):
#     """Fallback remark parsing without OpenAI."""
#     su = str(remark).upper()
#     fid, fn  = parse_fuel_from_text(remark)
#     iw, fw, tw_ = parse_weight(remark)
#     iss, fs, ts  = parse_seating(remark)
#     ii, fi, ti   = parse_idv(remark)
#     return {
#         'vehicle_make':    '', 'vehicle_model': '', 'vehicle_variant': '',
#         'is_with_ncb':     ncb_flag(su),
#         'is_irda_tp':      irda_flag(su),
#         'is_cpa_included': cpa_flag(su),
#         'is_nil_dep':      nil_dep_flag(su),
#         'fuel_type':       fn,
#         'seating_cap':     -1, 'from_seating': fs, 'to_seating': ts,
#         'is_weight':       iw, 'from_weight_kg': fw, 'to_weight_kg': tw_,
#         'idv_cap':         ii, 'from_idv': fi, 'to_idv': ti,
#         'is_cc':           -1, 'from_cc': 0, 'to_cc': 99999,
#     }

# _api_calls  = 0
# _cache_hits = 0
# _api_errors = 0
# _api_ms     = 0

# def parse_remark_openai(remark, co_name, seg, pt, row_n=0):
#     """
#     Call OpenAI to extract structured fields from remark.
#     Falls back to heuristic on error or missing key.
#     """
#     global _api_calls, _api_errors, _api_ms
#     ak = os.getenv("OPENAI_API_KEY", "").strip()
#     if not ak:
#         return heuristic_parse(remark, seg, pt)

#     included = keep_included_only(remark)
#     log("API", f"Row {row_n:>4} | call #{_api_calls+1} | {remark[:70]!r}")

#     prompt = f"""You are an expert Indian motor insurance data extractor.
# Analyse the remark carefully. Return ONLY valid JSON — no markdown, no extra text — with EXACTLY these keys:

#   vehicle_make    : INCLUDED vehicle makes, comma-separated. Empty string if none.
#   vehicle_model   : INCLUDED vehicle models, comma-separated. Empty string if none.
#   vehicle_variant : Specific variant if mentioned, else "".
#   is_with_ncb     : 1=NCB included, 0=WITHOUT/NON/ZERO NCB, -1=not mentioned.
#   is_irda_tp      : 1=IRDA TP rate mentioned, -1=otherwise.
#   is_cpa_included : 1=CPA included/mentioned, -1=otherwise.
#   is_nil_dep      : 1=Nil Dep / Zero Dep mentioned, -1=otherwise.
#   fuel_type       : DIESEL | PETROL | ELECTRIC | CNG-LPG | "" (empty if not mentioned).
#   seating_cap     : exact integer if single SC value, -1 otherwise.
#   from_seating    : lower SC bound (int), -1 if N/A.
#   to_seating      : upper SC bound (int), -1 if N/A.
#   is_weight       : 1=GVW/weight/tonnage mentioned, -1=otherwise.
#   from_weight_kg  : lower weight in KG (int), 0 if N/A.
#   to_weight_kg    : upper weight in KG (int), 99999 if N/A.
#   idv_cap         : 1=IDV cap mentioned, -1=otherwise.
#   from_idv        : lower IDV in Lacs (float), 0 if N/A.
#   to_idv          : upper IDV in Lacs (float), 0 if N/A.
#   is_cc           : 1=engine CC/capacity mentioned, -1=otherwise.
#   from_cc         : lower CC (int), 0 if N/A.
#   to_cc           : upper CC (int), 99999 if N/A.

# RULES:
# 1. ONLY list makes/models that are INCLUDED. IGNORE anything after EXCEPT/BUT/EXCLUDE/OTHER THAN/REJECT/DECLINE.
# 2. If remark is purely exclusion (e.g. "Except TATA" or "HR 68 EXCLUDED"), set vehicle_make="" vehicle_model="".
# 3. SC=Seating Capacity. "SC 7"→seating_cap=7. "SC upto 7"→from_seating=1 to_seating=7. "SC(3+1)"→seating_cap=4.
# 4. Tons→KG: 1T=1000KG. "7.5T"→to_weight_kg=7500.
# 5. "IDV upto 10 lacs"→idv_cap=1 from_idv=0 to_idv=10.
# 6. "Upto 1500 CC"→is_cc=1 from_cc=0 to_cc=1500.
# 7. TRACTOR segment→vehicle_variant="Agriculture Tractor".
# 8. Multiple makes/models separated by comma (e.g. "HONDA,HYUNDAI,KIA").

# company: {co_name}
# segment: {seg}
# policy_type: {pt}
# remark_original: {remark}
# remark_included_only: {included}
# """

#     body = {
#         "model": _ai_model,
#         "messages": [
#             {"role": "system", "content": "Return only valid JSON, no markdown, no explanation."},
#             {"role": "user",   "content": prompt}
#         ],
#         "response_format": {"type": "json_object"},
#         "temperature": 0
#     }
#     req = urllib.request.Request(
#         "https://api.openai.com/v1/chat/completions",
#         data=json.dumps(body).encode(),
#         headers={"Content-Type": "application/json",
#                  "Authorization": f"Bearer {ak}"},
#         method="POST"
#     )
#     ts = time.time()
#     try:
#         with urllib.request.urlopen(req, timeout=30) as r:
#             raw = json.loads(r.read())
#         p = json.loads(raw["choices"][0]["message"]["content"])
#         ms = int((time.time() - ts) * 1000)
#         _api_calls += 1; _api_ms += ms

#         def _s(k, d=""): return str(p.get(k, d)).strip()
#         def _i(k, d=-1):
#             try: return int(p.get(k, d))
#             except: return d
#         def _f(k, d=0.0):
#             try: return float(p.get(k, d))
#             except: return d

#         result = {
#             'vehicle_make':    _s('vehicle_make'),
#             'vehicle_model':   _s('vehicle_model'),
#             'vehicle_variant': _s('vehicle_variant'),
#             'is_with_ncb':     _i('is_with_ncb'),
#             'is_irda_tp':      _i('is_irda_tp'),
#             'is_cpa_included': _i('is_cpa_included'),
#             'is_nil_dep':      _i('is_nil_dep'),
#             'fuel_type':       _s('fuel_type').upper(),
#             'seating_cap':     _i('seating_cap'),
#             'from_seating':    _i('from_seating'),
#             'to_seating':      _i('to_seating'),
#             'is_weight':       _i('is_weight'),
#             'from_weight_kg':  _i('from_weight_kg', 0),
#             'to_weight_kg':    _i('to_weight_kg', 99999),
#             'idv_cap':         _i('idv_cap'),
#             'from_idv':        _f('from_idv'),
#             'to_idv':          _f('to_idv'),
#             'is_cc':           _i('is_cc'),
#             'from_cc':         _i('from_cc', 0),
#             'to_cc':           _i('to_cc', 99999),
#         }
#         log("OK", (f"Row {row_n:>4} | {ms}ms | "
#                    f"make={result['vehicle_make']!r} model={result['vehicle_model']!r} "
#                    f"ncb={result['is_with_ncb']} fuel={result['fuel_type']!r} cc={result['is_cc']}"))
#         return result

#     except urllib.error.HTTPError as e:
#         _api_errors += 1
#         log("ERR", f"Row {row_n:>4} | HTTP {e.code} → heuristic fallback")
#     except Exception as e:
#         _api_errors += 1
#         log("ERR", f"Row {row_n:>4} | {e} → heuristic fallback")
#     return heuristic_parse(remark, seg, pt)

# # =============================================================================
# #  OUTPUT ROW BUILDER
# # =============================================================================
# OUTPUT_COLUMNS = [
#     'id', 'company_id', 'company_code', 'segment_id', 'segment',
#     'subproduct_id', 'sub_product_name', 'lob_id', 'lob_name',
#     'business_type_id', 'business_type', 'is_highend_lob',
#     'rto_group_id', 'rto_group_name',
#     'payin_od_rate', 'payin_tp_rate', 'payout_od_rate', 'payout_tp_rate',
#     'extra_tp_rate', 'eff_from_date', 'eff_to_date',
#     'fuel_type_id', 'fuel_type',
#     'is_on_net', 'is_one_year_pay_on_newbusiness', 'is_cpa_included',
#     'is_geared_vehicle', 'is_cc_considered', 'from_cc', 'to_cc',
#     'is_premium_considered', 'from_premium', 'to_premium',
#     'is_mmv_considered', 'make_id', 'vehicle_make', 'model_id', 'vehicle_model',
#     'variant_id', 'vehicle_variant',
#     'is_seating_cap_consider', 'from_seating_cap', 'to_seating_cap',
#     'is_no_of_wheel_consider', 'from_no_of_wheel', 'to_no_of_wheel',
#     'vehicle_type_id', 'ppi_in', 'ppi_out',
#     'is_irda_tp_included', 'is_longterm_renewal_pay',
#     'is_weightage_considered', 'from_weightage_kg', 'to_weightage_kg',
#     'is_nil_dep_considered', 'is_organization_type',
#     'from_age_month', 'to_age_month',
#     'is_with_ncb', 'is_idv_cap_consider', 'from_idv', 'to_idv',
#     'is_breakin_consider', 'is_active',
# ]

# def build_row(
#     seg_id, seg_name, sp_id, sp_name, rto_id, rto_name,
#     pod, ptp, pod2, ptp2,
#     fuel_type_id, fuel_type_name, is_on_net, cpa, geared,
#     is_cc, from_cc, to_cc,
#     is_wt, from_wt, to_wt,
#     vt_id,
#     is_mmv, mk_id, mk_name, md_id, md_name, vr_id, vr_name,
#     is_sc, from_sc, to_sc,
#     is_whl, from_whl, to_whl,
#     from_age, to_age,
#     ncb, irda,
#     is_idv, from_idv, to_idv,
#     lob_name='',
#     nil_dep=-1,
# ):
#     return {
#         'id':                           0,
#         'company_id':                   CID,
#         'company_code':                 CCODE,
#         'segment_id':                   seg_id,
#         'segment':                      seg_name,
#         'subproduct_id':                sp_id,
#         'sub_product_name':             sp_name,
#         'lob_id':                       -1,
#         'lob_name':                     lob_name,
#         'business_type_id':             -1,
#         'business_type':                'Not Considered',
#         'is_highend_lob':               False,
#         'rto_group_id':                 rto_id,
#         'rto_group_name':               rto_name,
#         'payin_od_rate':                pod,
#         'payin_tp_rate':                ptp,
#         'payout_od_rate':               pod2,
#         'payout_tp_rate':               ptp2,
#         'extra_tp_rate':                0,
#         'eff_from_date':                '2026-01-01',
#         'eff_to_date':                  '2026-01-16',
#         'fuel_type_id':                 fuel_type_id,
#         'fuel_type':                    fuel_type_name,
#         'is_on_net':                    is_on_net,
#         'is_one_year_pay_on_newbusiness': -1,
#         'is_cpa_included':              cpa,
#         'is_geared_vehicle':            geared,
#         'is_cc_considered':             is_cc,
#         'from_cc':                      from_cc,
#         'to_cc':                        to_cc,
#         'is_premium_considered':        -1,
#         'from_premium':                 -1,
#         'to_premium':                   -1,
#         'is_mmv_considered':            is_mmv,
#         'make_id':                      mk_id,
#         'vehicle_make':                 mk_name,
#         'model_id':                     md_id,
#         'vehicle_model':                md_name,
#         'variant_id':                   vr_id,
#         'vehicle_variant':              vr_name,
#         'is_seating_cap_consider':      is_sc,
#         'from_seating_cap':             from_sc,
#         'to_seating_cap':               to_sc,
#         'is_no_of_wheel_consider':      is_whl,
#         'from_no_of_wheel':             from_whl,
#         'to_no_of_wheel':               to_whl,
#         'vehicle_type_id':              vt_id,
#         'ppi_in':                       0,
#         'ppi_out':                      0,
#         'is_irda_tp_included':          irda,
#         'is_longterm_renewal_pay':      -1,
#         'is_weightage_considered':      is_wt,
#         'from_weightage_kg':            from_wt,
#         'to_weightage_kg':              to_wt,
#         'is_nil_dep_considered':        nil_dep,
#         'is_organization_type':         -1,
#         'from_age_month':               from_age,
#         'to_age_month':                 to_age,
#         'is_with_ncb':                  ncb,
#         'is_idv_cap_consider':          is_idv,
#         'from_idv':                     from_idv,
#         'to_idv':                       to_idv,
#         'is_breakin_consider':          -1,
#         'is_active':                    True,
#     }

# # =============================================================================
# #  BUILD LOB NAME
# # =============================================================================

# def build_lob_name(co_name, seg_text, pt, remark):
#     """
#     Build a descriptive lob_name string.
#     Format: <company> <segment> <policy_type> <included_remark>
#     Extra info from remark (e.g. NCB cases, add-on cover conditions) goes here.
#     """
#     parts = [
#         str(co_name).strip(),
#         str(seg_text).strip(),
#         str(pt).strip(),
#         keep_included_only(remark).strip(),
#     ]
#     return ' '.join(p for p in parts if p)

# # =============================================================================
# #  PROCESS ONE INPUT FILE
# # =============================================================================

# def process_file(input_path):
#     global _cache_hits
#     log("INFO", f"Reading: {input_path}")
#     t_read = time.time()
#     df = pd.read_excel(input_path)
#     df.columns = [c.strip() for c in df.columns]
#     log("OK", f"Loaded {len(df)} rows in {time.time()-t_read:.1f}s | cols: {list(df.columns)}")

#     def col(*names):
#         """Return first matching column name that exists in df."""
#         for n in names:
#             if n in df.columns:
#                 return n
#         return None

#     # Column aliases
#     c_seg  = col('SEGMENT', 'Segment', 'LOB')
#     c_pt   = col('POLICY TYPE', 'Policy Type', 'POLICYTYPE')
#     c_loc  = col('LOCATION', 'Location', 'GEO LOCATION', 'Geo Location')
#     c_pay  = col('PAYIN', 'Payin', 'PAYIN (OD)', 'Payin (OD Premium)')
#     c_pout = col('PAYOUT', 'Payout', 'Calculated Payout', 'CALCULATED PAYOUT')
#     c_rem  = col('REMARK', 'Remark', 'REMARKS', 'Remarks', 'CALCULATION EXPLANATION')
#     c_age  = col('AGE', 'Age', 'AGE BAND', 'Age Band', 'AGE (YEARS)')
#     c_cc   = col('CC BAND', 'CC Band', 'CC', 'CC_BAND')
#     c_tw   = col('TW TYPE', 'TW Type', 'TW_TYPE')
#     c_co   = col('COMPANY NAME', 'Company Name', 'COMPANY', 'Company')

#     log("INFO", (f"Column map → seg={c_seg!r} pt={c_pt!r} loc={c_loc!r} "
#                  f"pay={c_pay!r} pout={c_pout!r} rem={c_rem!r} age={c_age!r} cc={c_cc!r}"))

#     # Count unique remarks
#     uniq_rem = set()
#     if c_rem:
#         uniq_rem = set(df[c_rem].fillna('').astype(str).str.strip())
#     log("INFO", f"Unique remarks: {len(uniq_rem)} → "
#         + (f"~{len(uniq_rem)} API calls" if _api_key else "heuristic only"))

#     out_rows = []
#     cache    = {}
#     total    = len(df)
#     t_start  = time.time()

#     print(f"\n  {'='*58}\n  Processing {total} rows …\n  {'='*58}\n")

#     for idx, (_, row) in enumerate(df.iterrows(), 1):
#         progress_bar(idx, total)

#         if idx == 1 or idx % 25 == 0 or idx == total:
#             el   = time.time() - t_start
#             rate = idx / el if el else 0
#             eta  = (total - idx) / rate if rate else 0
#             avg  = (_api_ms / _api_calls) if _api_calls else 0
#             log("INFO", (f"Row {idx:>4}/{total} | {el:.0f}s | ETA {eta:.0f}s | "
#                          f"rate={rate:.1f}/s | api={_api_calls} avg={avg:.0f}ms | "
#                          f"cache={_cache_hits} err={_api_errors} | out={len(out_rows)}"))

#         def g(c, d=''):
#             """Get cell value safely, return d if missing/null."""
#             if c is None:
#                 return d
#             v = row.get(c, d)
#             if v is None or str(v).strip() in ('nan', 'None', 'NaN', ''):
#                 return d
#             return v

#         # ── Policy Type ──────────────────────────────────────────────────────
#         pt      = normalize_policy_type(str(g(c_pt, 'COMP')).strip())
#         seg_nm  = POLICY_TO_SEG.get(pt, 'Comprehensive')
#         seg_id  = seg_name_to_id.get(seg_nm, 1)
#         is_on_net = (pt == 'COMP')

#         # ── Pay-in / Pay-out ─────────────────────────────────────────────────
#         payin  = sf(g(c_pay, 0))
#         payout = sf(g(c_pout, 0))
#         if   pt == 'TP':   pod, ptp, pod2, ptp2 = 0, payin, 0, payout
#         elif pt == 'SAOD': pod, ptp, pod2, ptp2 = payin, 0, payout, 0
#         else:              pod = ptp = payin; pod2 = ptp2 = payout

#         # ── Vehicle / Subproduct ─────────────────────────────────────────────
#         seg_text = str(g(c_seg, '')).strip()
#         sp_name, sp_id, vt_id = vehicle_info_from_segment(seg_text)
#         is_cv = sp_name in ('Goods Vehicle', 'Passenger Vehicle', 'Miscellaneous Vehicle')

#         # ── Location ─────────────────────────────────────────────────────────
#         rto_name = str(g(c_loc, '')).strip()
#         rto_id   = 0  # ID resolution done post-process if needed

#         # ── Scalar fields ────────────────────────────────────────────────────
#         remark   = str(g(c_rem, '')).strip()
#         co_name  = str(g(c_co, CCODE)).strip()
#         from_age, to_age  = parse_age(str(g(c_age, '')))
#         fcc0, tcc0, iscc0 = parse_cc_band(str(g(c_cc, '')))

#         # ── TW geared / vehicle type override ───────────────────────────────
#         tw_raw = str(g(c_tw, '')).strip().lower()
#         geared = -1
#         if sp_name == 'Two Wheeler' and tw_raw:
#             if 'scooter' in tw_raw:
#                 vt_id  = vt_name_to_id.get('TW Scooter', vt_id)
#                 geared = 0
#             elif 'bike' in tw_raw:
#                 vt_id  = vt_name_to_id.get('TW Bike', vt_id)
#                 geared = 1
#             elif 'electric' in tw_raw:
#                 vt_id  = vt_name_to_id.get('TW Electric Bike', vt_id)
#                 geared = -1

#         # ── Remark parsing (cached) ──────────────────────────────────────────
#         ck = (remark, co_name, seg_text, pt)
#         if ck in cache:
#             _cache_hits += 1
#             log("CACHE", f"Row {idx:>4} | HIT #{_cache_hits} | {remark[:50]!r}")
#             meta = cache[ck]
#         else:
#             meta       = parse_remark_openai(remark, co_name, seg_text, pt, idx)
#             cache[ck]  = meta

#         # ── Extract meta fields ──────────────────────────────────────────────
#         ncb    = meta['is_with_ncb']
#         irda   = meta['is_irda_tp']
#         cpa    = meta['is_cpa_included']
#         nil_dep= meta.get('is_nil_dep', -1)
#         raw_mk = meta['vehicle_make']
#         raw_md = meta['vehicle_model']
#         raw_vr = meta['vehicle_variant']

#         # Fuel
#         m_fuel = meta.get('fuel_type', '')
#         if m_fuel and m_fuel in fuel_id:
#             fid_v, fn_v = fuel_id[m_fuel], m_fuel
#         else:
#             fid_v, fn_v = parse_fuel_from_text(remark)

#         # Seating
#         msc, mfsc, mtsc = meta.get('seating_cap', -1), meta.get('from_seating', -1), meta.get('to_seating', -1)
#         if msc != -1:
#             issc_v, fsc_v, tsc_v = 1, msc, msc
#         elif mfsc != -1 or mtsc != -1:
#             issc_v = 1
#             fsc_v  = mfsc if mfsc != -1 else 1
#             tsc_v  = mtsc if mtsc != -1 else 99
#         else:
#             issc_v, fsc_v, tsc_v = parse_seating(remark)

#         # Weight
#         if meta.get('is_weight', -1) == 1:
#             iswt_v, fwt_v, twt_v = 1, meta['from_weight_kg'], meta['to_weight_kg']
#         else:
#             iswt_v, fwt_v, twt_v = parse_weight(remark)

#         # IDV
#         if meta.get('idv_cap', -1) == 1:
#             idv_v, fidv_v, tidv_v = 1, meta['from_idv'], meta['to_idv']
#         else:
#             idv_v, fidv_v, tidv_v = parse_idv(remark)

#         # CC (column takes priority over OpenAI)
#         iscc_v, fcc_v, tcc_v = iscc0, fcc0, tcc0
#         if iscc_v == -1 and meta.get('is_cc', -1) == 1:
#             iscc_v, fcc_v, tcc_v = 1, meta['from_cc'], meta['to_cc']

#         # Pure exclusion → clear MMV fields
#         if is_pure_exclusion(remark):
#             raw_mk = raw_md = raw_vr = ''
#             log("INFO", f"Row {idx:>4} | pure-exclusion remark → MMV cleared")

#         # LOB name
#         lob_name = build_lob_name(co_name, seg_text, pt, remark)

#         # ── Expand multiple makes (one output row per make) ──────────────────
#         make_list = [x.strip() for x in re.split(r'[,&]+', raw_mk) if x.strip()] or ['']
#         if len(make_list) > 1:
#             log("INFO", f"Row {idx:>4} | expanding {len(make_list)} makes: {make_list}")

#         for one_make in make_list:
#             # MMV resolution
#             if is_cv:
#                 mk_id, mk_name, md_id, md_name, vr_id, vr_name, i_fuel, i_seat = \
#                     resolve_cv_mmv(one_make, raw_md, raw_vr, remark)
#                 i_cc = i_gear = -1
#             else:
#                 mk_id, mk_name, md_id, md_name, vr_id, vr_name, i_fuel, i_seat, i_cc, i_gear = \
#                     resolve_tw_mmv(one_make, raw_md, raw_vr, remark)

#             if one_make and mk_id == -1:
#                 log("WARN", f"Row {idx:>4} | make '{one_make}' NOT found in MMV for {CCODE}")

#             # Tractor → variant override
#             if sp_name == 'Miscellaneous Vehicle' and 'TRACTOR' in seg_text.upper():
#                 if not vr_name:
#                     vr_name = 'Agriculture Tractor'

#             is_mmv = 1 if (mk_id != -1 or md_id != -1 or vr_id != -1 or one_make or raw_md or raw_vr) else -1

#             # Fuel fallback from MMV variant
#             fin_fid, fin_fn = fid_v, fn_v
#             if fin_fid == -1 and i_fuel and i_fuel in fuel_id:
#                 fin_fid, fin_fn = fuel_id[i_fuel], i_fuel

#             # Seating fallback from MMV
#             fin_issc, fin_fsc, fin_tsc = issc_v, fsc_v, tsc_v
#             if fin_issc == -1 and isinstance(i_seat, int) and i_seat > 0:
#                 fin_issc, fin_fsc, fin_tsc = 1, i_seat, i_seat

#             # Geared fallback from MMV
#             fin_gear = geared
#             if sp_name == 'Two Wheeler' and fin_gear == -1 and i_gear != -1:
#                 fin_gear = i_gear

#             out_rows.append(build_row(
#                 seg_id, seg_nm, sp_id, sp_name, rto_id, rto_name,
#                 pod, ptp, pod2, ptp2,
#                 fin_fid, fin_fn, is_on_net, cpa, fin_gear,
#                 iscc_v, fcc_v, tcc_v,
#                 iswt_v, fwt_v, twt_v,
#                 vt_id,
#                 is_mmv,
#                 mk_id,  mk_name  if mk_name  else one_make,
#                 md_id,  md_name  if md_name  else raw_md,
#                 vr_id,  vr_name  if vr_name  else raw_vr,
#                 fin_issc, fin_fsc, fin_tsc,
#                 -1, -1, -1,
#                 from_age, to_age,
#                 ncb, irda,
#                 idv_v, fidv_v, tidv_v,
#                 lob_name=lob_name,
#                 nil_dep=nil_dep,
#             ))

#     el  = time.time() - t_start
#     avg = (_api_ms / _api_calls) if _api_calls else 0
#     print()
#     log("OK", "=" * 60)
#     log("OK", f"DONE  input={total}  output={len(out_rows)}  time={el:.1f}s ({el/60:.1f}min)")
#     log("OK", f"  API calls={_api_calls}  avg={avg:.0f}ms  cache={_cache_hits}  errors={_api_errors}")
#     log("OK", "=" * 60)
#     return pd.DataFrame(out_rows)

# # =============================================================================
# #  MAIN LOOP
# # =============================================================================
# input_file = input("\nEnter input Excel file path: ").strip().strip('"')

# while True:
#     try:
#         df_out = process_file(input_file)
#         df_out = df_out[[c for c in OUTPUT_COLUMNS if c in df_out.columns]]

#         if os.path.exists(OUT_FILE):
#             log("INFO", f"Appending to existing output: {OUT_FILE}")
#             existing = pd.read_excel(OUT_FILE)
#             df_out   = pd.concat([existing, df_out], ignore_index=True)

#         df_out.to_excel(OUT_FILE, index=False)
#         log("OK", f"Saved → {OUT_FILE}  ({len(df_out)} rows)")
#         log("OK", f"Log   → {_log_path}")

#     except Exception as e:
#         import traceback
#         log("ERR", f"FATAL: {e}")
#         traceback.print_exc()

#     print("\n" + "=" * 75)
#     print("  1  Process another file (appends to same output)")
#     print("  2  Exit")
#     choice = input("Choice: ").strip()
#     if choice == "2":
#         log("OK", "Goodbye!")
#         if _log_file:
#             _log_file.close()
#         break
#     input_file = input("Next input file path: ").strip().strip('"')


"""
PayinConfig Generator — Universal Motor Insurance Pay-in Config Tool
Supports Shriram and any other company in the master list.

Features:
  - Loads master data from Masters_-_dec_2025.xlsx (Excel sheets)
  - OpenAI-powered remark parsing with heuristic fallback
  - Make/Model/Variant resolution for TW/PC (global) and CV (per company)
  - Dominant-ID strategy for CV MMV to avoid duplicate MakeID issues
  - Full column schema as per spec
  - Progress bar + colored logging + log file
  - Append mode: multiple input files -> single output
  - Remark caching to avoid duplicate API calls

Usage:
  python PayinConfig.py
  Set OPENAI_API_KEY in environment or in a .env file in the JSON folder.
"""

import pandas as  pd
import re, sys, os, json, time, urllib.request, urllib.error
from datetime import datetime

# =============================================================================
#  LOGGING
# =============================================================================
COLORS = {
    "OK"   : "\033[32m",
    "WARN" : "\033[33m",
    "ERR"  : "\033[31m",
    "API"  : "\033[36m",
    "CACHE": "\033[90m",
    "INFO" : "",
}
_log_file = None

def log(level, msg):
    ts   = datetime.now().strftime("%H:%M:%S")
    line = f"[{ts}][{level}] {msg}"
    col  = COLORS.get(level, "")
    print(f"{col}{line}\033[0m" if col else line)
    if _log_file:
        _log_file.write(line + "\n")
        _log_file.flush()

def progress_bar(cur, tot, width=50):
    pct = cur / tot if tot else 0
    filled = int(width * pct)
    bar = "█" * filled + "░" * (width - filled)
    print(f"\r  [{bar}] {cur}/{tot} ({pct*100:.1f}%)", end="", flush=True)
    if cur == tot:
        print()

# =============================================================================
#  UTILITIES
# =============================================================================

def nt(v):
    """Normalize: UPPER-CASE and collapse non-alphanumeric to single space."""
    return re.sub(r'\s+', ' ', re.sub(r'[^A-Z0-9]+', ' ', str(v).upper())).strip()

def si(v, default=-1):
    try:   return int(float(str(v)))
    except: return default

def sf(v, default=0.0):
    if isinstance(v, str):
        v = v.strip().replace('%', '')
    try:   return float(v)
    except: return default

def load_dotenv(path):
    if not os.path.exists(path):
        return
    with open(path, encoding="utf-8") as f:
        for line in f:
            s = line.strip()
            if not s or s.startswith('#') or '=' not in s:
                continue
            k, v = s.split('=', 1)
            k = k.strip()
            v = v.strip().strip('"').strip("'")
            if k and k not in os.environ:
                os.environ[k] = v

# =============================================================================
#  STARTUP
# =============================================================================
print("\n" + "="*75)
print("  PayinConfig Generator  — Motor Insurance Pay-in Config Tool")
print("="*75)

MASTERS  = input("Path to Masters_-_dec_2025.xlsx : ").strip().strip('"')
JSON_DIR = input("Path to JSON files folder       : ").strip().strip('"')
OUT_DIR  = input("Output folder path              : ").strip().strip('"')
os.makedirs(OUT_DIR, exist_ok=True)

_log_path = os.path.join(OUT_DIR, f"payin_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
try:
    _log_file = open(_log_path, "w", encoding="utf-8")
    log("OK", f"Log → {_log_path}")
except Exception as e:
    print(f"[WARN] Cannot open log file: {e}")

load_dotenv(os.path.join(JSON_DIR, ".env"))
load_dotenv(".env")

# =============================================================================
#  LOAD MASTER DATA FROM EXCEL
# =============================================================================
log("INFO", f"Loading master data: {MASTERS}")
t0 = time.time()
xl = pd.ExcelFile(MASTERS)

def read_sheet(name):
    df = pd.read_excel(xl, sheet_name=name)
    df.columns = [str(c).strip() for c in df.columns]
    return df

df_company = read_sheet("Insurance Company Master")
df_segment = read_sheet("Segment Master")
df_subprod = read_sheet("SubProduct Master")
df_vtype   = read_sheet("Vehicle type")
df_fuel    = read_sheet("fuel master")
df_tw_mmv  = read_sheet("tw_pc_mmv master")
df_cv_mmv  = read_sheet("cv_mmv master")

log("OK", f"Masters loaded in {time.time()-t0:.1f}s")

# ── Company lookup ──────────────────────────────────────────────────────────
company_by_id = {si(r['company_id']): r for _, r in df_company.iterrows()}

# ── Segment lookup ──────────────────────────────────────────────────────────
seg_name_to_id = {str(r['segment']).strip(): si(r['id'])
                  for _, r in df_segment.iterrows()}

POLICY_TO_SEG = {
    'COMP': 'Comprehensive',  'COMPREHENSIVE': 'Comprehensive',
    'TP':   'TP Only',        'TP ONLY':       'TP Only',   'THIRD PARTY': 'TP Only',
    'SAOD': 'SAOD',           'OD':            'SAOD',      'OWN DAMAGE':  'SAOD',
    'TW NEW': 'Comprehensive',
}

# ── Subproduct lookup (motor only) ─────────────────────────────────────────
if 'product_name' in df_subprod.columns:
    df_sp = df_subprod[df_subprod['product_name'].str.upper() == 'MOTOR']
else:
    df_sp = df_subprod
subprod_id = {str(r['sub_product_name']).strip(): si(r['sub_product_id'])
              for _, r in df_sp.iterrows()}

# ── Vehicle type lookup ─────────────────────────────────────────────────────
vt_name_to_id  = {str(r['vehicle_type']).strip(): si(r['id'])
                  for _, r in df_vtype.iterrows()}
vt_sub_default = {}  # sub_product_name → first vehicle_type_id
for _, r in df_vtype.iterrows():
    sp = str(r['sub_product_name']).strip()
    if sp not in vt_sub_default:
        vt_sub_default[sp] = si(r['id'])

# ── Fuel lookup ─────────────────────────────────────────────────────────────
fuel_id = {str(r['fuel_type']).strip().upper(): si(r['id'])
           for _, r in df_fuel.iterrows()}

_FUEL_ALIASES = {
    'PETROL': 'PETROL',   'P': 'PETROL',
    'DIESEL': 'DIESEL',   'D': 'DIESEL',
    'ELECTRIC': 'ELECTRIC', 'E': 'ELECTRIC', 'EV': 'ELECTRIC', 'B': 'ELECTRIC',
    'CNG-LPG': 'CNG-LPG', 'CNG': 'CNG-LPG', 'LPG': 'CNG-LPG', 'C': 'CNG-LPG',
    'CNG/PETROL': 'CNG-LPG', 'HYBRID': 'PETROL', 'PETROL/ELECTRIC': 'ELECTRIC',
}

def fuel_lookup(raw):
    """Return (fuel_type_id, fuel_type_name) from raw string."""
    k = _FUEL_ALIASES.get(str(raw).strip().upper(), '')
    return fuel_id.get(k, -1), k

# =============================================================================
#  MMV INDEX BUILD
# =============================================================================
log("INFO", "Building MMV indexes …")
t1 = time.time()

# ── TW / Private Car MMV (global, no company scope) ──────────────────────────
tw_make           = {}  # norm_make → (make_id, display_make)
tw_model          = {}  # (norm_make, norm_model) → (model_id, display_model)
tw_var            = {}  # (norm_make, norm_model, norm_var) → (var_id, display_var, fuel, seat, cc, geared)
tw_makes_sorted   = []  # [(norm, display, id)] sorted by len(norm) desc
tw_models_of_make = {}  # norm_make → set of (norm_model, display_model, model_id)

for _, m in df_tw_mmv.iterrows():
    mkn = nt(m.get('vehicle_make', ''))
    mdn = nt(m.get('vehicle_model', ''))
    vrn = nt(m.get('VariantDisplayName', ''))
    mid  = si(m.get('MakeID', -1))
    mdid = si(m.get('ModelID', -1))
    vid  = si(m.get('VariantID', -1))
    fid_v, fn = fuel_lookup(m.get('Fuel', ''))
    seat = si(m.get('SeatingCapacity', -1))
    cc   = si(m.get('CC', -1))
    gear = 1 if str(m.get('Is_Geared_Vehicle', '')).upper() in ('TRUE', '1', 'YES') else 0

    if mkn and mid != -1:
        tw_make.setdefault(mkn, (mid, str(m.get('vehicle_make', '')).strip()))
    if mkn and mdn and mdid != -1:
        tw_model.setdefault((mkn, mdn), (mdid, str(m.get('vehicle_model', '')).strip()))
        tw_models_of_make.setdefault(mkn, set()).add(
            (mdn, str(m.get('vehicle_model', '')).strip(), mdid))
    if mkn and mdn and vrn and vid != -1:
        tw_var.setdefault((mkn, mdn, vrn),
                          (vid, str(m.get('VariantDisplayName', '')).strip(), fn, seat, cc, gear))

tw_makes_sorted = sorted(
    [(k, tw_make[k][1], tw_make[k][0]) for k in tw_make],
    key=lambda x: len(x[0]), reverse=True)

# ── CV MMV (per company code) ─────────────────────────────────────────────────
# Pre-compute dominant MakeID/ModelID to handle duplicate batches in master
_cv_make_cnt = {}   # (code, make_norm, make_id) → row count
_cv_model_cnt = {}  # (code, make_norm, model_norm, model_id) → row count

for _, m in df_cv_mmv.iterrows():
    code = nt(str(m.get('company_code', '')))
    mkn  = nt(str(m.get('vehicle_make', '')))
    mdn  = nt(str(m.get('vehicle_model', '')))
    mid  = si(m.get('MakeID', -1))
    mdid = si(m.get('ModelID', -1))
    if code and mkn and mid != -1:
        _cv_make_cnt[(code, mkn, mid)] = _cv_make_cnt.get((code, mkn, mid), 0) + 1
    if code and mkn and mdn and mdid != -1:
        _cv_model_cnt[(code, mkn, mdn, mdid)] = _cv_model_cnt.get((code, mkn, mdn, mdid), 0) + 1

# For each (code, make_norm) choose MakeID with most rows
_dom_make_id  = {}
_make_agg = {}
for (code, mkn, mid), cnt in _cv_make_cnt.items():
    _make_agg.setdefault((code, mkn), {})[mid] = cnt
for (code, mkn), id_cnt in _make_agg.items():
    _dom_make_id[(code, mkn)] = max(id_cnt, key=id_cnt.get)

# For each (code, make_norm, model_norm) choose ModelID with most rows
_dom_model_id = {}
_model_agg = {}
for (code, mkn, mdn, mdid), cnt in _cv_model_cnt.items():
    _model_agg.setdefault((code, mkn, mdn), {})[mdid] = cnt
for (code, mkn, mdn), id_cnt in _model_agg.items():
    _dom_model_id[(code, mkn, mdn)] = max(id_cnt, key=id_cnt.get)

cv_make          = {}  # (code, norm_make) → (make_id, display_make)
cv_model         = {}  # (code, norm_make, norm_model) → (model_id, display_model)
cv_var           = {}  # (code, norm_make, norm_model, norm_var) → (var_id, display_var, fuel, seat)
cv_makes_of_co   = {}  # code → sorted [(norm, display, id)]
cv_models_of_mk  = {}  # (code, norm_make) → set of (norm_model, display_model, model_id)
cv_models_of_co  = {}  # code → {norm_model: [(model_id, display_model, make_id, display_make)]}

for _, m in df_cv_mmv.iterrows():
    code = nt(m.get('company_code', ''))
    mkn  = nt(m.get('vehicle_make', ''))
    mdn  = nt(m.get('vehicle_model', ''))
    vrn  = nt(m.get('VariantName', ''))
    mid  = si(m.get('MakeID', -1))
    mdid = si(m.get('ModelID', -1))
    vid  = si(m.get('VariantID', -1))
    fid_v, fn = fuel_lookup(m.get('Fuel', ''))
    seat = si(m.get('SeatingCapacity', -1))
    mk_disp = str(m.get('vehicle_make', '')).strip()
    md_disp = str(m.get('vehicle_model', '')).strip()

    if code and mkn and mid != -1:
        dom_mid = _dom_make_id.get((code, mkn), mid)
        if mid == dom_mid or (code, mkn) not in cv_make:
            cv_make[(code, mkn)] = (dom_mid, mk_disp)
        cv_makes_of_co.setdefault(code, set()).add((mkn, mk_disp, dom_mid))

    if code and mkn and mdn and mdid != -1:
        dom_mid  = _dom_make_id.get((code, mkn), mid)
        dom_mdid = _dom_model_id.get((code, mkn, mdn), mdid)
        if mdid == dom_mdid or (code, mkn, mdn) not in cv_model:
            cv_model[(code, mkn, mdn)] = (dom_mdid, md_disp)
        cv_models_of_mk.setdefault((code, mkn), set()).add((mdn, md_disp, dom_mdid))
        cv_models_of_co.setdefault(code, {}).setdefault(mdn, [])
        existing = {x[0] for x in cv_models_of_co[code][mdn]}
        if dom_mdid not in existing:
            cv_models_of_co[code][mdn].append((dom_mdid, md_disp, dom_mid, mk_disp))

    if code and mkn and mdn and vrn and vid != -1:
        cv_var.setdefault((code, mkn, mdn, vrn),
                          (vid, str(m.get('VariantName', '')).strip(), fn, seat))

for k in cv_makes_of_co:
    cv_makes_of_co[k] = sorted(list(cv_makes_of_co[k]),
                                key=lambda x: len(x[0]), reverse=True)

log("OK", (f"MMV indexes built in {time.time()-t1:.1f}s | "
           f"TW makes={len(tw_make)} models={len(tw_model)} variants={len(tw_var)} | "
           f"CV makes={len(cv_make)} models={len(cv_model)} variants={len(cv_var)}"))

# =============================================================================
#  COMPANY SELECTION
# =============================================================================
print("\n" + "="*75)
print("AVAILABLE COMPANIES")
print("="*75)
for cid, row in sorted(company_by_id.items()):
    print(f"  {cid:3d}  |  {str(row['company_code']):20s}  |  {row['company_name']}")
print("="*75)

while True:
    try:
        CID      = int(input("\nEnter company_id: ").strip())
        CROW     = company_by_id[CID]
        CCODE    = str(CROW['company_code']).strip()
        CCODE_NT = nt(CCODE)
        log("OK", f"Company → {CROW['company_name']}  (id={CID}  code={CCODE})")
        break
    except (ValueError, KeyError):
        log("ERR", "Invalid company_id — please try again.")

OUT_FILE = os.path.join(OUT_DIR, f"{CCODE}-Payin-Config.xlsx")

_api_key = os.getenv("OPENAI_API_KEY", "").strip()
_ai_model = os.getenv("OPENAI_MODEL", "gpt-4.1-mini")
if _api_key:
    log("OK",   f"OpenAI key found — model: {_ai_model}")
else:
    log("WARN", "No OPENAI_API_KEY — heuristic remark parsing only")

# =============================================================================
#  PARSING HELPERS
# =============================================================================

def normalize_policy_type(raw):
    s = str(raw).strip().upper()
    if s in ('COMP', 'COMPREHENSIVE'):           return 'COMP'
    if s in ('TP', 'TP ONLY', 'THIRD PARTY'):   return 'TP'
    if s in ('SAOD', 'OD', 'OWN DAMAGE'):       return 'SAOD'
    if s == 'TW NEW':                            return 'COMP'
    return 'COMP'

def vehicle_info_from_segment(seg_text):
    """
    Infer (sub_product_name, sub_product_id, vehicle_type_id) from SEGMENT column.
    Handles TW, PC, PCV, GCV, CV, Misc, Tractor, etc.
    """
    s = str(seg_text).strip().upper()

    if re.search(r'\bTW\b', s) or re.search(r'\b2W\b', s):
        sp = 'Two Wheeler'
        return sp, subprod_id.get(sp, -1), vt_name_to_id.get('TW Bike', vt_sub_default.get(sp, -1))

    if re.search(r'\bPVT\s*CAR\b', s) or re.search(r'\bPRIVATE\s*CAR\b', s) or re.search(r'\bPC\b', s):
        sp = 'Private Car'
        return sp, subprod_id.get(sp, -1), vt_name_to_id.get('Private Car', vt_sub_default.get(sp, -1))

    if re.search(r'\bPCV\b', s) or re.search(r'\bPASSENGER\b', s):
        sp = 'Passenger Vehicle'
        return sp, subprod_id.get(sp, -1), vt_name_to_id.get('Auto rikshaw', vt_sub_default.get(sp, -1))

    if 'TRACTOR' in s or 'HARVESTER' in s:
        sp = 'Miscellaneous Vehicle'
        return sp, subprod_id.get(sp, -1), vt_name_to_id.get('Agriculture Tractor', vt_sub_default.get(sp, -1))

    if 'MISD' in s or 'MISC' in s:
        sp = 'Miscellaneous Vehicle'
        return sp, subprod_id.get(sp, -1), vt_name_to_id.get('Non Tractor', vt_sub_default.get(sp, -1))

    # Default: Goods Vehicle (GCV / GVW / CV)
    sp = 'Goods Vehicle'
    return sp, subprod_id.get(sp, -1), vt_name_to_id.get('Truck', vt_sub_default.get(sp, -1))

def parse_age(s):
    """Parse age field → (from_age_month, to_age_month)."""
    s = str(s).strip()
    if s.lower() in ('', 'nan', 'none', 'new', 'n/a'):
        return 0, 700
    su = s.upper()

    m = re.search(r'UPTO\s+(\d+)\s*(YEAR|YR|MONTH|MTH|MO)', su)
    if m:
        n, u = int(m.group(1)), m.group(2)
        return 0, n * 12 if u.startswith('Y') else n

    m = re.search(r'(\d+)\s*TO\s*(\d+)\s*(YEAR|YR)?', su)
    if m:
        return int(m.group(1)) * 12, int(m.group(2)) * 12

    m = re.match(r'>\s*(\d+)\s*[-]\s*(\d+)(\+?)\s*[Yy]', s)
    if m:
        lo = int(m.group(1)) * 12 + 1
        hi = 700 if m.group(3) == '+' else int(m.group(2)) * 12
        return lo, hi

    m = re.match(r'>\s*(\d+)\+?\s*[Yy]', s)
    if m:
        return int(m.group(1)) * 12 + 1, 700

    m = re.match(r'^(\d+)$', s.strip())
    if m:
        return 0, int(m.group(1)) * 12

    return 0, 700

def parse_cc_band(s):
    """Parse CC band field → (from_cc, to_cc, is_cc_considered)."""
    if not s or str(s).strip().lower() in ('', 'nan', 'none'):
        return 0, 99999, -1
    s = str(s).strip().upper().replace('CC', '').strip()
    patterns = [
        (r'^<\s*(\d+)$',                 lambda m: (0, int(m.group(1)) - 1, 1)),
        (r'^>\s*(\d+)$',                 lambda m: (int(m.group(1)) + 1, 99999, 1)),
        (r'^>\s*(\d+)\s*[-]\s*(\d+)$',  lambda m: (int(m.group(1)) + 1, int(m.group(2)), 1)),
        (r'^>=\s*(\d+)\s*[-]\s*(\d+)$', lambda m: (int(m.group(1)), int(m.group(2)), 1)),
        (r'^(\d+)\s*[-]\s*(\d+)$',      lambda m: (int(m.group(1)), int(m.group(2)), 1)),
        (r'^(\d+)$',                     lambda m: (int(m.group(1)), int(m.group(1)), 1)),
    ]
    for pat, fn in patterns:
        mt = re.match(pat, s)
        if mt:
            return fn(mt)
    return 0, 99999, -1

def parse_idv(text):
    """Parse IDV cap from remark → (is_idv_cap, from_idv, to_idv)."""
    su = str(text).upper()
    m = re.search(r'IDV\s+(?:UPTO|UP\s*TO|<=?)\s*([\d.]+)\s*(?:LAC|LAKH|L\b)', su)
    if m:
        return 1, 0.0, float(m.group(1))
    m = re.search(r'IDV\s+([\d.]+)\s*[-]\s*([\d.]+)\s*(?:LAC|LAKH)', su)
    if m:
        return 1, float(m.group(1)), float(m.group(2))
    return -1, 0.0, 0.0

def parse_weight(text):
    """Parse weight/tonnage from remark → (is_weight, from_kg, to_kg)."""
    su = str(text).upper()
    m = re.search(r'([\d.]+)\s*KG\b', su)
    if m:
        return 1, 0, int(float(m.group(1)))
    m = re.search(r'UPTO\s+([\d.]+)\s*T(?:ON|ONNE)?', su)
    if m:
        return 1, 0, int(float(m.group(1)) * 1000)
    m = re.search(r'>\s*([\d.]+)\s*[-]\s*([\d.]+)\s*T(?:ON|ONNE)?', su)
    if m:
        return 1, int(float(m.group(1)) * 1000) + 1, int(float(m.group(2)) * 1000)
    m = re.search(r'\b([\d.]+)\s*T(?:ON|ONNE)?\b', su)
    if m:
        return 1, 0, int(float(m.group(1)) * 1000)
    m = re.search(r'GVW\s+([\d.]+)', su)
    if m:
        v = float(m.group(1))
        return 1, 0, int(v * 1000) if v < 200 else int(v)
    return -1, 0, 99999

def parse_seating(text):
    """Parse seating capacity from remark → (is_seating, from_sc, to_sc)."""
    su = str(text).upper()
    m = re.search(r'(?:SC|SEATING\s*CAPACITY|SEATING)\s*UPTO\s*(\d+)', su)
    if m:
        return 1, 1, int(m.group(1))
    m = re.search(r'(?:SC|SEATING\s*CAPACITY|SEATING)\s*(\d+)\s*[-]\s*(\d+)', su)
    if m:
        return 1, int(m.group(1)), int(m.group(2))
    m = re.search(r'(?:SC|SEATING\s*CAPACITY|SEATING)\s*\(?\s*(\d+)', su)
    if m:
        v = int(m.group(1))
        return 1, v, v
    m = re.search(r'(\d+)\s*SEATER', su)
    if m:
        v = int(m.group(1))
        return 1, v, v
    return -1, -1, -1

def parse_fuel_from_text(text):
    """Detect fuel type from free text → (fuel_type_id, fuel_type_name)."""
    su = str(text).upper()
    if 'ELECTRIC' in su or ' EV ' in su:
        return fuel_id.get('ELECTRIC', -1), 'ELECTRIC'
    if 'CNG' in su or 'LPG' in su:
        return fuel_id.get('CNG-LPG', -1), 'CNG-LPG'
    if 'DIESEL' in su:
        return fuel_id.get('DIESEL', -1), 'DIESEL'
    if 'PETROL' in su:
        return fuel_id.get('PETROL', -1), 'PETROL'
    return -1, ''

def ncb_flag(text):
    su = str(text).upper()
    if any(x in su for x in ('WITHOUT NCB', 'W/O NCB', 'NON NCB', 'NON-NCB', 'ZERO NCB')):
        return 0
    if 'NCB' in su:
        return 1
    return -1

def irda_flag(text):
    su = str(text).upper()
    return 1 if any(x in su for x in ('IRDA TP', 'IRDA RATE', 'IRDA')) else -1

def cpa_flag(text):
    su = str(text).upper()
    if 'CPA' in su and any(x in su for x in ('INCLUD', 'WITH CPA')):
        return 1
    return -1

def nil_dep_flag(text):
    su = str(text).upper()
    if 'NIL DEP' in su or 'NIL DEPRECIATION' in su or 'ZERO DEP' in su:
        return 1
    return -1

def keep_included_only(text):
    """Strip excluded/rejected clauses; return only the included portion."""
    s = str(text).strip()
    if not s or s.lower() in ('nan', 'none'):
        return ''
    kept = []
    for chunk in re.split(r';', s):
        up = chunk.upper()
        if any(t in up for t in ('DECLIN', 'REJECT')):
            continue
        for tok in (' BUT ', ' EXCEPT ', ' EXCLUDE ', ' OTHER THAN ', ' NOT CONSIDER'):
            idx = up.find(tok)
            if idx != -1:
                chunk = chunk[:idx]
                break
        chunk = chunk.strip()
        if chunk:
            kept.append(chunk)
    return ' '.join(x for x in kept if x)

def is_pure_exclusion(text):
    """Return True if remark contains ONLY exclusion language."""
    s = str(text).strip().upper()
    if not s or s in ('NAN', 'NONE', ''):
        return False
    EXCL_TOKENS = ('EXCEPT', 'EXCLUDE', 'OTHER THAN', 'REJECT', 'DECLIN',
                   'NOT CONSIDER', 'HR 68', 'EXCLUDED')
    INCL_TOKENS = ('ONLY', 'INCLUD', 'HONDA', 'BAJAJ', 'TATA', 'MARUTI',
                   'HYUNDAI', 'KIA', 'MAHINDRA', 'IDV', 'NCB', 'ZONE',
                   'BRANCH', 'NCB CASES')
    if any(t in s for t in INCL_TOKENS):
        return False
    return any(t in s for t in EXCL_TOKENS)

# =============================================================================
#  MMV RESOLUTION
# =============================================================================

def _score_match(query, candidate):
    """
    Score how well candidate matches query (both normalized).
    Returns (score, tie_breaker). Higher = better.
    100: exact  80: all words + starts with  60: words subset
    40: query in candidate  20: candidate in query  -1: no match
    """
    if not query or not candidate:
        return -1, 0
    qw = set(query.split())
    cw = set(candidate.split())
    extra = len(cw) - len(qw)

    if query == candidate:
        return 100, -extra
    if qw.issubset(cw) and candidate.startswith(query):
        return 80, -extra
    if qw.issubset(cw):
        return 60, -extra
    if query in candidate:
        return 40, -extra
    if candidate in query:
        return 20, -extra
    return -1, 0

def _best_match(query_norm, candidates):
    """
    candidates: list of (norm, display, id)
    Returns (id, display, matched_norm) or (-1, '', '')
    """
    best_score, best_tie, best = -1, 0, None
    for cn, cname, cid in candidates:
        score, tie = _score_match(query_norm, cn)
        if score > best_score or (score == best_score and tie > best_tie):
            best_score, best_tie, best = score, tie, (cid, cname, cn)
    if best_score >= 20:
        return best
    return -1, '', ''


def resolve_tw_mmv(raw_make, raw_model, raw_variant, remark):
    """
    Resolve TW/PC make/model/variant from MMV master.
    Returns (make_id, make_name, model_id, model_name, var_id, var_name,
             fuel_name, seat, cc, geared)
    """
    mkn = nt(raw_make) if raw_make else ''
    mdn = nt(raw_model) if raw_model else ''
    vrn = nt(raw_variant) if raw_variant else ''
    mk_id = mk_name = -1, ''
    md_id = md_name = -1, ''
    vr_id = vr_name = -1, ''
    fuel = ''; seat = cc = gear = -1
    mk_id, mk_name = -1, ''
    md_id, md_name = -1, ''
    vr_id, vr_name = -1, ''
    inc = nt(keep_included_only(remark))

    # 1. Make
    if mkn:
        if mkn in tw_make:
            mk_id, mk_name = tw_make[mkn]
        else:
            mk_id, mk_name, mkn = _best_match(mkn, tw_makes_sorted)
        if mk_id == -1:
            for cn, cname, cid in tw_makes_sorted:
                if cn and re.search(rf'\b{re.escape(cn)}\b', inc):
                    mk_id, mk_name, mkn = cid, cname, cn
                    break

    # 2. Model
    if mdn and mk_id != -1:
        if (mkn, mdn) in tw_model:
            md_id, md_name = tw_model[(mkn, mdn)]
        else:
            cands = [(mn, mname, mid) for mn, mname, mid in tw_models_of_make.get(mkn, set())]
            md_id, md_name, mdn = _best_match(mdn, cands)

    # 3. Variant
    if vrn and mk_id != -1 and md_id != -1:
        key = (mkn, mdn, vrn)
        if key in tw_var:
            vr_id, vr_name, fuel, seat, cc, gear = tw_var[key]
        else:
            for (cmn, cmdn, cvn), (vid, vname, vf, vs, vc, vg) in tw_var.items():
                if cmn == mkn and cmdn == mdn and (vrn in cvn or cvn in vrn):
                    vr_id, vr_name, fuel, seat, cc, gear = vid, vname, vf, vs, vc, vg
                    break

    # Infer fuel/seat/gear from first variant when no variant specified
    if mk_id != -1 and md_id != -1 and not vrn:
        for (cmn, cmdn, _), (vid, vname, vf, vs, vc, vg) in tw_var.items():
            if cmn == mkn and cmdn == mdn and vf:
                fuel, seat, cc, gear = vf, vs, vc, vg
                break

    return mk_id, mk_name, md_id, md_name, vr_id, vr_name, fuel, seat, cc, gear


def resolve_cv_mmv(raw_make, raw_model, raw_variant, remark):
    """
    Resolve CV make/model/variant from MMV master (scoped by CCODE_NT).
    Returns (make_id, make_name, model_id, model_name, var_id, var_name, fuel, seat)
    """
    code = CCODE_NT
    mkn  = nt(raw_make) if raw_make else ''
    mdn  = nt(raw_model) if raw_model else ''
    vrn  = nt(raw_variant) if raw_variant else ''
    mk_id, mk_name = -1, ''
    md_id, md_name = -1, ''
    vr_id, vr_name = -1, ''
    fuel = ''; seat = -1
    inc  = nt(keep_included_only(remark))

    # 1. Make
    if mkn:
        if (code, mkn) in cv_make:
            mk_id, mk_name = cv_make[(code, mkn)]
        else:
            mk_id, mk_name, mkn = _best_match(mkn, cv_makes_of_co.get(code, []))
        if mk_id == -1:
            for cn, cname, cid in cv_makes_of_co.get(code, []):
                if cn and re.search(rf'\b{re.escape(cn)}\b', inc):
                    mk_id, mk_name, mkn = cid, cname, cn
                    break

    # 2. Model
    if mdn:
        if mk_id != -1 and (code, mkn, mdn) in cv_model:
            md_id, md_name = cv_model[(code, mkn, mdn)]
        elif mk_id != -1:
            cands = [(mn, mname, mid) for mn, mname, mid in cv_models_of_mk.get((code, mkn), set())]
            md_id, md_name, mdn = _best_match(mdn, cands)
        if md_id == -1:
            hits = cv_models_of_co.get(code, {}).get(mdn, [])
            if hits:
                md_id, md_name, mk_id, mk_name = hits[0]

    # 3. Variant
    if vrn and mk_id != -1 and md_id != -1:
        key = (code, mkn, mdn, vrn)
        if key in cv_var:
            vr_id, vr_name, fuel, seat = cv_var[key]
        else:
            for (ck, cmn, cmdn, cvn), (vid, vname, vf, vs) in cv_var.items():
                if ck == code and cmn == mkn and cmdn == mdn and (vrn in cvn or cvn in vrn):
                    vr_id, vr_name, fuel, seat = vid, vname, vf, vs
                    break

    if mk_id != -1 and md_id != -1 and not vrn:
        for (ck, cmn, cmdn, _), (vid, vname, vf, vs) in cv_var.items():
            if ck == code and cmn == mkn and cmdn == mdn and vf:
                fuel, seat = vf, vs
                break

    return mk_id, mk_name, md_id, md_name, vr_id, vr_name, fuel, seat

# =============================================================================
#  OPENAI REMARK PARSER
# =============================================================================

def heuristic_parse(remark, seg, pt):
    """Fallback remark parsing without OpenAI."""
    su = str(remark).upper()
    fid, fn  = parse_fuel_from_text(remark)
    iw, fw, tw_ = parse_weight(remark)
    iss, fs, ts  = parse_seating(remark)
    ii, fi, ti   = parse_idv(remark)
    return {
        'vehicle_make':    '', 'vehicle_model': '', 'vehicle_variant': '',
        'is_with_ncb':     ncb_flag(su),
        'is_irda_tp':      irda_flag(su),
        'is_cpa_included': cpa_flag(su),
        'is_nil_dep':      nil_dep_flag(su),
        'fuel_type':       fn,
        'seating_cap':     -1, 'from_seating': fs, 'to_seating': ts,
        'is_weight':       iw, 'from_weight_kg': fw, 'to_weight_kg': tw_,
        'idv_cap':         ii, 'from_idv': fi, 'to_idv': ti,
        'is_cc':           -1, 'from_cc': 0, 'to_cc': 99999,
    }

_api_calls  = 0
_cache_hits = 0
_api_errors = 0
_api_ms     = 0

def parse_remark_openai(remark, co_name, seg, pt, row_n=0):
    """
    Call OpenAI to extract structured fields from remark.
    Falls back to heuristic on error or missing key.
    """
    global _api_calls, _api_errors, _api_ms
    ak = os.getenv("OPENAI_API_KEY", "").strip()
    if not ak:
        return heuristic_parse(remark, seg, pt)

    included = keep_included_only(remark)
    log("API", f"Row {row_n:>4} | call #{_api_calls+1} | {remark[:70]!r}")

    prompt = f"""You are an expert Indian motor insurance data extractor.
Analyse the remark carefully. Return ONLY valid JSON — no markdown, no extra text — with EXACTLY these keys:

  vehicle_make    : INCLUDED vehicle makes, comma-separated. Empty string if none.
  vehicle_model   : INCLUDED vehicle models, comma-separated. Empty string if none.
  vehicle_variant : Specific variant if mentioned, else "".
  is_with_ncb     : 1=NCB included, 0=WITHOUT/NON/ZERO NCB, -1=not mentioned.
  is_irda_tp      : 1=IRDA TP rate mentioned, -1=otherwise.
  is_cpa_included : 1=CPA included/mentioned, -1=otherwise.
  is_nil_dep      : 1=Nil Dep / Zero Dep mentioned, -1=otherwise.
  fuel_type       : DIESEL | PETROL | ELECTRIC | CNG-LPG | "" (empty if not mentioned).
  seating_cap     : exact integer if single SC value, -1 otherwise.
  from_seating    : lower SC bound (int), -1 if N/A.
  to_seating      : upper SC bound (int), -1 if N/A.
  is_weight       : 1=GVW/weight/tonnage mentioned, -1=otherwise.
  from_weight_kg  : lower weight in KG (int), 0 if N/A.
  to_weight_kg    : upper weight in KG (int), 99999 if N/A.
  idv_cap         : 1=IDV cap mentioned, -1=otherwise.
  from_idv        : lower IDV in Lacs (float), 0 if N/A.
  to_idv          : upper IDV in Lacs (float), 0 if N/A.
  is_cc           : 1=engine CC/capacity mentioned, -1=otherwise.
  from_cc         : lower CC (int), 0 if N/A.
  to_cc           : upper CC (int), 99999 if N/A.

RULES:
1. ONLY list makes/models that are INCLUDED. IGNORE anything after EXCEPT/BUT/EXCLUDE/OTHER THAN/REJECT/DECLINE.
2. If remark is purely exclusion (e.g. "Except TATA" or "HR 68 EXCLUDED"), set vehicle_make="" vehicle_model="".
3. SC=Seating Capacity. "SC 7"→seating_cap=7. "SC upto 7"→from_seating=1 to_seating=7. "SC(3+1)"→seating_cap=4.
4. Tons→KG: 1T=1000KG. "7.5T"→to_weight_kg=7500.
5. "IDV upto 10 lacs"→idv_cap=1 from_idv=0 to_idv=10.
6. "Upto 1500 CC"→is_cc=1 from_cc=0 to_cc=1500.
7. TRACTOR segment→vehicle_variant="Agriculture Tractor".
8. Multiple makes/models separated by comma (e.g. "HONDA,HYUNDAI,KIA").
Note: Sometimes , taxi can be said as kaali peeli taxi or kaali peeli ambassador taxi , note here is that ambasaador is the vehicle model and kaali peeli is just the attributes , so please consider ambassador as the vehicle model , dont write the kaali peeli
company: {co_name}
segment: {seg}
policy_type: {pt}
remark_original: {remark}
remark_included_only: {included}
"""

    body = {
        "model": _ai_model,
        "messages": [
            {"role": "system", "content": "Return only valid JSON, no markdown, no explanation."},
            {"role": "user",   "content": prompt}
        ],
        "response_format": {"type": "json_object"},
        "temperature": 0
    }
    req = urllib.request.Request(
        "https://api.openai.com/v1/chat/completions",
        data=json.dumps(body).encode(),
        headers={"Content-Type": "application/json",
                 "Authorization": f"Bearer {ak}"},
        method="POST"
    )
    ts = time.time()
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            raw = json.loads(r.read())
        p = json.loads(raw["choices"][0]["message"]["content"])
        ms = int((time.time() - ts) * 1000)
        _api_calls += 1; _api_ms += ms

        def _s(k, d=""): return str(p.get(k, d)).strip()
        def _i(k, d=-1):
            try: return int(p.get(k, d))
            except: return d
        def _f(k, d=0.0):
            try: return float(p.get(k, d))
            except: return d

        result = {
            'vehicle_make':    _s('vehicle_make'),
            'vehicle_model':   _s('vehicle_model'),
            'vehicle_variant': _s('vehicle_variant'),
            'is_with_ncb':     _i('is_with_ncb'),
            'is_irda_tp':      _i('is_irda_tp'),
            'is_cpa_included': _i('is_cpa_included'),
            'is_nil_dep':      _i('is_nil_dep'),
            'fuel_type':       _s('fuel_type').upper(),
            'seating_cap':     _i('seating_cap'),
            'from_seating':    _i('from_seating'),
            'to_seating':      _i('to_seating'),
            'is_weight':       _i('is_weight'),
            'from_weight_kg':  _i('from_weight_kg', 0),
            'to_weight_kg':    _i('to_weight_kg', 99999),
            'idv_cap':         _i('idv_cap'),
            'from_idv':        _f('from_idv'),
            'to_idv':          _f('to_idv'),
            'is_cc':           _i('is_cc'),
            'from_cc':         _i('from_cc', 0),
            'to_cc':           _i('to_cc', 99999),
        }
        log("OK", (f"Row {row_n:>4} | {ms}ms | "
                   f"make={result['vehicle_make']!r} model={result['vehicle_model']!r} "
                   f"ncb={result['is_with_ncb']} fuel={result['fuel_type']!r} cc={result['is_cc']}"))
        return result

    except urllib.error.HTTPError as e:
        _api_errors += 1
        log("ERR", f"Row {row_n:>4} | HTTP {e.code} → heuristic fallback")
    except Exception as e:
        _api_errors += 1
        log("ERR", f"Row {row_n:>4} | {e} → heuristic fallback")
    return heuristic_parse(remark, seg, pt)

# =============================================================================
#  OUTPUT ROW BUILDER
# =============================================================================
OUTPUT_COLUMNS = [
    'id', 'company_id', 'company_code', 'segment_id', 'segment',
    'subproduct_id', 'sub_product_name', 'lob_id', 'lob_name',
    'business_type_id', 'business_type', 'is_highend_lob',
    'rto_group_id', 'rto_group_name',
    'payin_od_rate', 'payin_tp_rate', 'payout_od_rate', 'payout_tp_rate',
    'extra_tp_rate', 'eff_from_date', 'eff_to_date',
    'fuel_type_id', 'fuel_type',
    'is_on_net', 'is_one_year_pay_on_newbusiness', 'is_cpa_included',
    'is_geared_vehicle', 'is_cc_considered', 'from_cc', 'to_cc',
    'is_premium_considered', 'from_premium', 'to_premium',
    'is_mmv_considered', 'make_id', 'vehicle_make', 'model_id', 'vehicle_model',
    'variant_id', 'vehicle_variant',
    'is_seating_cap_consider', 'from_seating_cap', 'to_seating_cap',
    'is_no_of_wheel_consider', 'from_no_of_wheel', 'to_no_of_wheel',
    'vehicle_type_id', 'ppi_in', 'ppi_out',
    'is_irda_tp_included', 'is_longterm_renewal_pay',
    'is_weightage_considered', 'from_weightage_kg', 'to_weightage_kg',
    'is_nil_dep_considered', 'is_organization_type',
    'from_age_month', 'to_age_month',
    'is_with_ncb', 'is_idv_cap_consider', 'from_idv', 'to_idv',
    'is_breakin_consider', 'is_active',
]

def build_row(
    seg_id, seg_name, sp_id, sp_name, rto_id, rto_name,
    pod, ptp, pod2, ptp2,
    fuel_type_id, fuel_type_name, is_on_net, cpa, geared,
    is_cc, from_cc, to_cc,
    is_wt, from_wt, to_wt,
    vt_id,
    is_mmv, mk_id, mk_name, md_id, md_name, vr_id, vr_name,
    is_sc, from_sc, to_sc,
    is_whl, from_whl, to_whl,
    from_age, to_age,
    ncb, irda,
    is_idv, from_idv, to_idv,
    lob_name='',
    nil_dep=-1,
):
    return {
        'id':                           0,
        'company_id':                   CID,
        'company_code':                 CCODE,
        'segment_id':                   seg_id,
        'segment':                      seg_name,
        'subproduct_id':                sp_id,
        'sub_product_name':             sp_name,
        'lob_id':                       -1,
        'lob_name':                     lob_name,
        'business_type_id':             -1,
        'business_type':                'Not Considered',
        'is_highend_lob':               False,
        'rto_group_id':                 rto_id,
        'rto_group_name':               rto_name,
        'payin_od_rate':                pod,
        'payin_tp_rate':                ptp,
        'payout_od_rate':               pod2,
        'payout_tp_rate':               ptp2,
        'extra_tp_rate':                0,
        'eff_from_date':                '2026-01-01',
        'eff_to_date':                  '2026-01-16',
        'fuel_type_id':                 fuel_type_id,
        'fuel_type':                    fuel_type_name,
        'is_on_net':                    is_on_net,
        'is_one_year_pay_on_newbusiness': -1,
        'is_cpa_included':              cpa,
        'is_geared_vehicle':            geared,
        'is_cc_considered':             is_cc,
        'from_cc':                      from_cc,
        'to_cc':                        to_cc,
        'is_premium_considered':        -1,
        'from_premium':                 -1,
        'to_premium':                   -1,
        'is_mmv_considered':            is_mmv,
        'make_id':                      mk_id,
        'vehicle_make':                 mk_name,
        'model_id':                     md_id,
        'vehicle_model':                md_name,
        'variant_id':                   vr_id,
        'vehicle_variant':              vr_name,
        'is_seating_cap_consider':      is_sc,
        'from_seating_cap':             from_sc,
        'to_seating_cap':               to_sc,
        'is_no_of_wheel_consider':      is_whl,
        'from_no_of_wheel':             from_whl,
        'to_no_of_wheel':               to_whl,
        'vehicle_type_id':              vt_id,
        'ppi_in':                       0,
        'ppi_out':                      0,
        'is_irda_tp_included':          irda,
        'is_longterm_renewal_pay':      -1,
        'is_weightage_considered':      is_wt,
        'from_weightage_kg':            from_wt,
        'to_weightage_kg':              to_wt,
        'is_nil_dep_considered':        nil_dep,
        'is_organization_type':         -1,
        'from_age_month':               from_age,
        'to_age_month':                 to_age,
        'is_with_ncb':                  ncb,
        'is_idv_cap_consider':          is_idv,
        'from_idv':                     from_idv,
        'to_idv':                       to_idv,
        'is_breakin_consider':          -1,
        'is_active':                    True,
    }

# =============================================================================
#  BUILD LOB NAME
# =============================================================================

def build_lob_name(co_name, seg_text, pt, remark):
    """
    Build a descriptive lob_name string.
    Format: <company> <segment> <policy_type> <included_remark>
    Extra info from remark (e.g. NCB cases, add-on cover conditions) goes here.
    """
    parts = [
        str(co_name).strip(),
        str(seg_text).strip(),
        str(pt).strip(),
        keep_included_only(remark).strip(),
    ]
    return ' '.join(p for p in parts if p)

# =============================================================================
#  PROCESS ONE INPUT FILE
# =============================================================================

def process_file(input_path):
    global _cache_hits
    log("INFO", f"Reading: {input_path}")
    t_read = time.time()
    df = pd.read_excel(input_path)
    df.columns = [c.strip() for c in df.columns]
    log("OK", f"Loaded {len(df)} rows in {time.time()-t_read:.1f}s | cols: {list(df.columns)}")

    def col(*names):
        """Return first matching column name that exists in df."""
        for n in names:
            if n in df.columns:
                return n
        return None

    # Column aliases
    c_seg  = col('SEGMENT', 'Segment', 'LOB')
    c_pt   = col('POLICY TYPE', 'Policy Type', 'POLICYTYPE')
    c_loc  = col('LOCATION', 'Location', 'GEO LOCATION', 'Geo Location')
    c_pay  = col('PAYIN', 'Payin', 'PAYIN (OD)', 'Payin (OD Premium)')
    c_pout = col('PAYOUT', 'Payout', 'Calculated Payout', 'CALCULATED PAYOUT')
    c_rem  = col('REMARK', 'Remark', 'REMARKS', 'Remarks', 'CALCULATION EXPLANATION')
    c_age  = col('AGE', 'Age', 'AGE BAND', 'Age Band', 'AGE (YEARS)')
    c_cc   = col('CC BAND', 'CC Band', 'CC', 'CC_BAND')
    c_tw   = col('TW TYPE', 'TW Type', 'TW_TYPE')
    c_co   = col('COMPANY NAME', 'Company Name', 'COMPANY', 'Company')

    log("INFO", (f"Column map → seg={c_seg!r} pt={c_pt!r} loc={c_loc!r} "
                 f"pay={c_pay!r} pout={c_pout!r} rem={c_rem!r} age={c_age!r} cc={c_cc!r}"))

    # Count unique remarks
    uniq_rem = set()
    if c_rem:
        uniq_rem = set(df[c_rem].fillna('').astype(str).str.strip())
    log("INFO", f"Unique remarks: {len(uniq_rem)} → "
        + (f"~{len(uniq_rem)} API calls" if _api_key else "heuristic only"))

    out_rows = []
    cache    = {}
    total    = len(df)
    t_start  = time.time()

    print(f"\n  {'='*58}\n  Processing {total} rows …\n  {'='*58}\n")

    for idx, (_, row) in enumerate(df.iterrows(), 1):
        progress_bar(idx, total)

        if idx == 1 or idx % 25 == 0 or idx == total:
            el   = time.time() - t_start
            rate = idx / el if el else 0
            eta  = (total - idx) / rate if rate else 0
            avg  = (_api_ms / _api_calls) if _api_calls else 0
            log("INFO", (f"Row {idx:>4}/{total} | {el:.0f}s | ETA {eta:.0f}s | "
                         f"rate={rate:.1f}/s | api={_api_calls} avg={avg:.0f}ms | "
                         f"cache={_cache_hits} err={_api_errors} | out={len(out_rows)}"))

        def g(c, d=''):
            """Get cell value safely, return d if missing/null."""
            if c is None:
                return d
            v = row.get(c, d)
            if v is None or str(v).strip() in ('nan', 'None', 'NaN', ''):
                return d
            return v

        # ── Policy Type ──────────────────────────────────────────────────────
        pt      = normalize_policy_type(str(g(c_pt, 'COMP')).strip())
        seg_nm  = POLICY_TO_SEG.get(pt, 'Comprehensive')
        seg_id  = seg_name_to_id.get(seg_nm, 1)
        is_on_net = (pt == 'COMP')

        # ── Pay-in / Pay-out ─────────────────────────────────────────────────
        payin  = sf(g(c_pay, 0))
        payout = sf(g(c_pout, 0))
        if   pt == 'TP':   pod, ptp, pod2, ptp2 = 0, payin, 0, payout
        elif pt == 'SAOD': pod, ptp, pod2, ptp2 = payin, 0, payout, 0
        else:              pod = ptp = payin; pod2 = ptp2 = payout

        # ── Vehicle / Subproduct ─────────────────────────────────────────────
        seg_text = str(g(c_seg, '')).strip()
        sp_name, sp_id, vt_id = vehicle_info_from_segment(seg_text)
        is_cv = sp_name in ('Goods Vehicle', 'Passenger Vehicle', 'Miscellaneous Vehicle')

        # ── Location ─────────────────────────────────────────────────────────
        rto_name = str(g(c_loc, '')).strip()
        rto_id   = 0  # ID resolution done post-process if needed

        # ── Scalar fields ────────────────────────────────────────────────────
        remark   = str(g(c_rem, '')).strip()
        co_name  = str(g(c_co, CCODE)).strip()
        from_age, to_age  = parse_age(str(g(c_age, '')))
        fcc0, tcc0, iscc0 = parse_cc_band(str(g(c_cc, '')))

        # ── TW geared / vehicle type override ───────────────────────────────
        tw_raw = str(g(c_tw, '')).strip().lower()
        geared = -1
        if sp_name == 'Two Wheeler' and tw_raw:
            if 'scooter' in tw_raw:
                vt_id  = vt_name_to_id.get('TW Scooter', vt_id)
                geared = 0
            elif 'bike' in tw_raw:
                vt_id  = vt_name_to_id.get('TW Bike', vt_id)
                geared = 1
            elif 'electric' in tw_raw:
                vt_id  = vt_name_to_id.get('TW Electric Bike', vt_id)
                geared = -1

        # ── Remark parsing (cached) ──────────────────────────────────────────
        ck = (remark, co_name, seg_text, pt)
        if ck in cache:
            _cache_hits += 1
            log("CACHE", f"Row {idx:>4} | HIT #{_cache_hits} | {remark[:50]!r}")
            meta = cache[ck]
        else:
            meta       = parse_remark_openai(remark, co_name, seg_text, pt, idx)
            cache[ck]  = meta

        # ── Extract meta fields ──────────────────────────────────────────────
        ncb    = meta['is_with_ncb']
        irda   = meta['is_irda_tp']
        cpa    = meta['is_cpa_included']
        nil_dep= meta.get('is_nil_dep', -1)
        raw_mk = meta['vehicle_make']
        raw_md = meta['vehicle_model']
        raw_vr = meta['vehicle_variant']

        # Fuel
        m_fuel = meta.get('fuel_type', '')
        if m_fuel and m_fuel in fuel_id:
            fid_v, fn_v = fuel_id[m_fuel], m_fuel
        else:
            fid_v, fn_v = parse_fuel_from_text(remark)

        # Seating
        msc, mfsc, mtsc = meta.get('seating_cap', -1), meta.get('from_seating', -1), meta.get('to_seating', -1)
        if msc != -1:
            issc_v, fsc_v, tsc_v = 1, msc, msc
        elif mfsc != -1 or mtsc != -1:
            issc_v = 1
            fsc_v  = mfsc if mfsc != -1 else 1
            tsc_v  = mtsc if mtsc != -1 else 99
        else:
            issc_v, fsc_v, tsc_v = parse_seating(remark)

        # Weight
        if meta.get('is_weight', -1) == 1:
            iswt_v, fwt_v, twt_v = 1, meta['from_weight_kg'], meta['to_weight_kg']
        else:
            iswt_v, fwt_v, twt_v = parse_weight(remark)

        # IDV
        if meta.get('idv_cap', -1) == 1:
            idv_v, fidv_v, tidv_v = 1, meta['from_idv'], meta['to_idv']
        else:
            idv_v, fidv_v, tidv_v = parse_idv(remark)

        # CC (column takes priority over OpenAI)
        iscc_v, fcc_v, tcc_v = iscc0, fcc0, tcc0
        if iscc_v == -1 and meta.get('is_cc', -1) == 1:
            iscc_v, fcc_v, tcc_v = 1, meta['from_cc'], meta['to_cc']

        # Pure exclusion → clear MMV fields
        if is_pure_exclusion(remark):
            raw_mk = raw_md = raw_vr = ''
            log("INFO", f"Row {idx:>4} | pure-exclusion remark → MMV cleared")

        # LOB name
        lob_name = build_lob_name(co_name, seg_text, pt, remark)

        # ── Expand multiple makes AND multiple models into separate rows ──────
        #
        # Rules:
        #   • If only models are listed (no make) → one row per model
        #   • If only makes are listed (no model)  → one row per make
        #   • If both makes AND models → one row per make (model shared),
        #     UNLESS models count > 1 AND makes count == 1, in which case
        #     one row per model under that make.
        #   • If both makes > 1 AND models > 1 → treat every token as either
        #     a make or a model via MMV lookup and expand individually.
        #
        # In practice: split BOTH fields on commas/ampersands/slashes and
        # build a flat list of (make_token, model_token) pairs to resolve.

        def _split_tokens(s):
            """Split a comma/ampersand/slash-separated string into clean tokens."""
            return [x.strip() for x in re.split(r'[,&/]+', s) if x.strip()]

        make_tokens  = _split_tokens(raw_mk)
        model_tokens = _split_tokens(raw_md)

        # Build list of (make_str, model_str) pairs to process
        if make_tokens and model_tokens:
            # Both present — if counts differ, pair each model with each make
            # (Cartesian product only when it makes sense, otherwise zip)
            if len(make_tokens) == 1:
                # One make, multiple models → one row per model
                pairs = [(make_tokens[0], m) for m in model_tokens]
            elif len(model_tokens) == 1:
                # Multiple makes, one model → one row per make
                pairs = [(mk, model_tokens[0]) for mk in make_tokens]
            else:
                # Multiple makes AND multiple models → one row per make,
                # pair by position (zip), leftover makes get empty model
                pairs = list(zip(make_tokens, model_tokens))
                # Append any extra makes/models that didn't get paired
                for mk in make_tokens[len(model_tokens):]:
                    pairs.append((mk, ''))
                for md in model_tokens[len(make_tokens):]:
                    pairs.append(('', md))
        elif make_tokens:
            pairs = [(mk, '') for mk in make_tokens]
        elif model_tokens:
            pairs = [('', md) for md in model_tokens]
        else:
            pairs = [('', '')]

        if len(pairs) > 1:
            log("INFO", f"Row {idx:>4} | expanding {len(pairs)} MMV pairs: {pairs}")

        def _emit_row(one_make, one_model):
            """Resolve MMV for one (make, model) pair and append to out_rows."""
            if is_cv:
                mk_id, mk_name, md_id, md_name, vr_id, vr_name, i_fuel, i_seat = \
                    resolve_cv_mmv(one_make, one_model, raw_vr, remark)
                i_cc = i_gear = -1
            else:
                mk_id, mk_name, md_id, md_name, vr_id, vr_name, i_fuel, i_seat, i_cc, i_gear = \
                    resolve_tw_mmv(one_make, one_model, raw_vr, remark)

            if one_make and mk_id == -1:
                log("WARN", f"Row {idx:>4} | make '{one_make}' NOT found in MMV for {CCODE}")
            if one_model and md_id == -1:
                log("WARN", f"Row {idx:>4} | model '{one_model}' NOT found in MMV for {CCODE}")

            # Tractor → variant override
            vr_name_final = vr_name
            if sp_name == 'Miscellaneous Vehicle' and 'TRACTOR' in seg_text.upper():
                if not vr_name_final:
                    vr_name_final = 'Agriculture Tractor'

            is_mmv = 1 if (
                mk_id != -1 or md_id != -1 or vr_id != -1
                or one_make or one_model or raw_vr
            ) else -1

            # Fuel fallback from MMV variant
            fin_fid, fin_fn = fid_v, fn_v
            if fin_fid == -1 and i_fuel and i_fuel in fuel_id:
                fin_fid, fin_fn = fuel_id[i_fuel], i_fuel

            # Seating fallback from MMV
            fin_issc, fin_fsc, fin_tsc = issc_v, fsc_v, tsc_v
            if fin_issc == -1 and isinstance(i_seat, int) and i_seat > 0:
                fin_issc, fin_fsc, fin_tsc = 1, i_seat, i_seat

            # Geared fallback from MMV
            fin_gear = geared
            if sp_name == 'Two Wheeler' and fin_gear == -1 and i_gear != -1:
                fin_gear = i_gear

            out_rows.append(build_row(
                seg_id, seg_nm, sp_id, sp_name, rto_id, rto_name,
                pod, ptp, pod2, ptp2,
                fin_fid, fin_fn, is_on_net, cpa, fin_gear,
                iscc_v, fcc_v, tcc_v,
                iswt_v, fwt_v, twt_v,
                vt_id,
                is_mmv,
                mk_id, mk_name if mk_name else one_make,
                md_id, md_name if md_name else one_model,
                vr_id, vr_name_final if vr_name_final else raw_vr,
                fin_issc, fin_fsc, fin_tsc,
                -1, -1, -1,
                from_age, to_age,
                ncb, irda,
                idv_v, fidv_v, tidv_v,
                lob_name=lob_name,
                nil_dep=nil_dep,
            ))

        for one_make, one_model in pairs:
            _emit_row(one_make, one_model)

    el  = time.time() - t_start
    avg = (_api_ms / _api_calls) if _api_calls else 0
    print()
    log("OK", "=" * 60)
    log("OK", f"DONE  input={total}  output={len(out_rows)}  time={el:.1f}s ({el/60:.1f}min)")
    log("OK", f"  API calls={_api_calls}  avg={avg:.0f}ms  cache={_cache_hits}  errors={_api_errors}")
    log("OK", "=" * 60)
    return pd.DataFrame(out_rows)

# =============================================================================#
#  MAIN LOOP                                                                   #
# =============================================================================#
input_file = input("\nEnter input Excel file path: ").strip().strip('"')

while True:
    try:
        df_out = process_file(input_file)
        df_out = df_out[[c for c in OUTPUT_COLUMNS if c in df_out.columns]]

        if os.path.exists(OUT_FILE):
            log("INFO", f"Appending to existing output: {OUT_FILE}")
            existing = pd.read_excel(OUT_FILE)
            df_out   = pd.concat([existing, df_out], ignore_index=True)

        df_out.to_excel(OUT_FILE, index=False)
        log("OK", f"Saved → {OUT_FILE}  ({len(df_out)} rows)")
        log("OK", f"Log   → {_log_path}")

    except Exception as e:
        import traceback
        log("ERR", f"FATAL: {e}")
        traceback.print_exc()

    print("\n" + "=" * 75)
    print("  1  Process another file (appends to same output)")
    print("  2  Exit")
    choice = input("Choice: ").strip()
    if choice == "2":
        log("OK", "Goodbye!")
        if _log_file:
            _log_file.close()
        break
    input_file = input("Next input file path: ").strip().strip('"')
