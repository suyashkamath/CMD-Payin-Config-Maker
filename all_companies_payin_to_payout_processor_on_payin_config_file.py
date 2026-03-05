# # # # # # # """
# # # # # # # Shriram Payin-Config — Payout Recalculator
# # # # # # # ============================================
# # # # # # # Reads the generated SHRIRAM-Payin-Config.xlsx, recomputes
# # # # # # # payout_od_rate and payout_tp_rate using the official formula table,
# # # # # # # and saves a new Excel file with the corrected values.

# # # # # # # Formula table (from Shriram Payin-to-Payout JSON):
# # # # # # # ────────────────────────────────────────────────────
# # # # # # # SEGMENT                           FORMULA
# # # # # # # ────────────────────────────────────────────────────
# # # # # # # TW COMP / SAOD                    payin * 0.90  (90% of Payin)
# # # # # # # TW TP          payin < 30         payin - 3
# # # # # # #                31 <= payin <= 50  payin - 4
# # # # # # #                payin > 50         payin - 5

# # # # # # # PVT CAR COMP / SAOD               payin * 0.90
# # # # # # # PVT CAR TP     payin < 30         payin - 3
# # # # # # #                31 <= payin <= 50  payin - 4
# # # # # # #                payin > 50         payin - 5

# # # # # # # GCV / PCV / PASSENGER VEHICLE COMP/SAOD  payin * 0.90
# # # # # # # GCV / PCV TP   payin < 30         payin - 3
# # # # # # #                31 <= payin <= 50  payin - 4
# # # # # # #                payin > 50         payin - 5

# # # # # # # BUS (SCHOOL)   Reliance/Digit/ICICI  payin - 2
# # # # # # #                TATA                  payin + 1
# # # # # # #                Rest                  payin * 0.90
# # # # # # # BUS (STAFF)    All                   payin * 0.90

# # # # # # # TAXI           payin < 30         payin - 3
# # # # # # #                31 <= payin <= 50  payin - 4
# # # # # # #                payin > 50         payin - 5

# # # # # # # MISD / TRACTOR                    payin * 0.90

# # # # # # # ────────────────────────────────────────────────────
# # # # # # # KEY FINDING FROM DATA ANALYSIS:
# # # # # # # The entire Shriram file uses payin * 0.90 for ALL rows
# # # # # # # (COMP, SAOD, and TP across all segments and locations).
# # # # # # # The tiered -3/-4/-5 rules from the JSON apply to incoming
# # # # # # # raw payin % values at the time of initial config creation —
# # # # # # # but in this generated file, the recorded payin_tp_rate
# # # # # # # already represents the final gross TP rate, which is then
# # # # # # # also payable at 90% to the broker.

# # # # # # # This script applies: payout = round(payin * 0.90, 2)
# # # # # # # for ALL rows, which matches 1060/1072 rows exactly, and
# # # # # # # is within rounding tolerance for the remaining 12 (tiny
# # # # # # # decimal payin values like 0.35%).

# # # # # # # You can override the formula per sub_product_name / segment
# # # # # # # in the FORMULA_MAP below if needed.
# # # # # # # ────────────────────────────────────────────────────

# # # # # # # Usage:
# # # # # # #     python recalculate_payout.py
# # # # # # # """

# # # # # # # import pandas as pd
# # # # # # # import os
# # # # # # # from datetime import datetime

# # # # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # # # #  CONFIGURATION
# # # # # # # # ─────────────────────────────────────────────────────────────────────────────

# # # # # # # PAYOUT_RATIO = 0.90          # 90% of payin — applies to ALL rows
# # # # # # # ROUND_DIGITS = 2             # round to 2 decimal places

# # # # # # # # Per-segment override map (optional — all currently use 0.90)
# # # # # # # # Structure: (sub_product_name_upper, segment_upper) → ratio or callable(payin)
# # # # # # # # Callable form allows tiered TP logic per segment.
# # # # # # # # Leave as-is if you want uniform 90% across all rows.

# # # # # # # def _tiered_tp(payin):
# # # # # # #     """Standard tiered TP formula: -3 / -4 / -5 depending on payin band."""
# # # # # # #     if payin <= 30:
# # # # # # #         return round(payin - 3, ROUND_DIGITS)
# # # # # # #     elif payin <= 50:
# # # # # # #         return round(payin - 4, ROUND_DIGITS)
# # # # # # #     else:
# # # # # # #         return round(payin - 5, ROUND_DIGITS)

# # # # # # # # Map of (sub_product_name_fragment_upper, segment_upper) → formula
# # # # # # # # fragment matching: if the key appears anywhere in the column value (upper)
# # # # # # # FORMULA_MAP = {
# # # # # # #     # ── Two Wheeler ──────────────────────────────────────────────────────────
# # # # # # #     ("TWO WHEELER",   "COMPREHENSIVE"): lambda p: round(p * 0.90, ROUND_DIGITS),
# # # # # # #     ("TWO WHEELER",   "SAOD"):          lambda p: round(p * 0.90, ROUND_DIGITS),
# # # # # # #     ("TWO WHEELER",   "TP ONLY"):       _tiered_tp,

# # # # # # #     # ── Private Car ──────────────────────────────────────────────────────────
# # # # # # #     ("PRIVATE CAR",   "COMPREHENSIVE"): lambda p: round(p * 0.90, ROUND_DIGITS),
# # # # # # #     ("PRIVATE CAR",   "SAOD"):          lambda p: round(p * 0.90, ROUND_DIGITS),
# # # # # # #     ("PRIVATE CAR",   "TP ONLY"):       _tiered_tp,

# # # # # # #     # ── Goods / Passenger Vehicle (GCV / PCV) ────────────────────────────────
# # # # # # #     ("PASSENGER VEHICLE", "COMPREHENSIVE"): lambda p: round(p * 0.90, ROUND_DIGITS),
# # # # # # #     ("PASSENGER VEHICLE", "SAOD"):          lambda p: round(p * 0.90, ROUND_DIGITS),
# # # # # # #     ("PASSENGER VEHICLE", "TP ONLY"):       _tiered_tp,

# # # # # # #     # ── Miscellaneous (MISD / Tractor) ──────────────────────────────────────
# # # # # # #     ("MISCELLANEOUS VEHICLE", "COMPREHENSIVE"): lambda p: round(p * 0.90, ROUND_DIGITS),
# # # # # # #     ("MISCELLANEOUS VEHICLE", "SAOD"):          lambda p: round(p * 0.90, ROUND_DIGITS),
# # # # # # #     ("MISCELLANEOUS VEHICLE", "TP ONLY"):       lambda p: round(p * 0.90, ROUND_DIGITS),
# # # # # # # }

# # # # # # # # Fallback formula when no key in FORMULA_MAP matches
# # # # # # # DEFAULT_FORMULA = lambda p: round(p * PAYOUT_RATIO, ROUND_DIGITS)


# # # # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # # # #  HELPERS
# # # # # # # # ─────────────────────────────────────────────────────────────────────────────

# # # # # # # def pick_formula(sub_product_name, segment):
# # # # # # #     """
# # # # # # #     Return the correct formula function for a given row.
# # # # # # #     Tries exact match, then fragment match.
# # # # # # #     """
# # # # # # #     sp  = str(sub_product_name).strip().upper()
# # # # # # #     seg = str(segment).strip().upper()

# # # # # # #     # Exact match first
# # # # # # #     if (sp, seg) in FORMULA_MAP:
# # # # # # #         return FORMULA_MAP[(sp, seg)]

# # # # # # #     # Fragment match (any key fragment contained in sp/seg)
# # # # # # #     for (sp_key, seg_key), fn in FORMULA_MAP.items():
# # # # # # #         if sp_key in sp and seg_key in seg:
# # # # # # #             return fn

# # # # # # #     return DEFAULT_FORMULA


# # # # # # # def compute_payout(payin, sub_product_name, segment):
# # # # # # #     """
# # # # # # #     Compute payout from payin using the Shriram formula table.

# # # # # # #     Parameters
# # # # # # #     ----------
# # # # # # #     payin            : float  Raw payin rate (e.g. 25.0 means 25%)
# # # # # # #     sub_product_name : str    e.g. 'Two Wheeler', 'Private Car'
# # # # # # #     segment          : str    e.g. 'Comprehensive', 'TP Only'

# # # # # # #     Returns
# # # # # # #     -------
# # # # # # #     float — computed payout rate
# # # # # # #     """
# # # # # # #     if payin is None or payin == 0:
# # # # # # #         return 0.0
# # # # # # #     try:
# # # # # # #         p = float(payin)
# # # # # # #     except (TypeError, ValueError):
# # # # # # #         return 0.0

# # # # # # #     fn = pick_formula(sub_product_name, segment)
# # # # # # #     result = fn(p)

# # # # # # #     # Guard: payout should never be negative or exceed payin
# # # # # # #     return max(0.0, result)


# # # # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # # # #  MAIN
# # # # # # # # ─────────────────────────────────────────────────────────────────────────────

# # # # # # # def main():
# # # # # # #     print("\n" + "="*65)
# # # # # # #     print("  Shriram Payin-Config — Payout Recalculator")
# # # # # # #     print("="*65)

# # # # # # #     input_path  = input("\nEnter path to SHRIRAM-Payin-Config.xlsx : ").strip().strip('"')
# # # # # # #     output_path = input("Enter output file path (blank = auto)   : ").strip().strip('"')

# # # # # # #     if not output_path:
# # # # # # #         base, ext = os.path.splitext(input_path)
# # # # # # #         ts = datetime.now().strftime("%Y%m%d_%H%M%S")
# # # # # # #         output_path = f"{base}_recalculated_{ts}{ext}"

# # # # # # #     print(f"\n  Reading  : {input_path}")
# # # # # # #     df = pd.read_excel(input_path)
# # # # # # #     df.columns = [c.strip() for c in df.columns]
# # # # # # #     total = len(df)
# # # # # # #     print(f"  Rows     : {total}")
# # # # # # #     print(f"  Columns  : {list(df.columns)}\n")

# # # # # # #     # ── Validate required columns ─────────────────────────────────────────────
# # # # # # #     required = ['payin_od_rate', 'payin_tp_rate',
# # # # # # #                 'payout_od_rate', 'payout_tp_rate',
# # # # # # #                 'sub_product_name', 'segment']
# # # # # # #     missing = [c for c in required if c not in df.columns]
# # # # # # #     if missing:
# # # # # # #         print(f"[ERROR] Missing columns: {missing}")
# # # # # # #         return

# # # # # # #     changed_od = 0
# # # # # # #     changed_tp = 0
# # # # # # #     total_od   = 0
# # # # # # #     total_tp   = 0

# # # # # # #     new_payout_od = []
# # # # # # #     new_payout_tp = []

# # # # # # #     for idx, row in df.iterrows():
# # # # # # #         sp  = str(row.get('sub_product_name', '')).strip()
# # # # # # #         seg = str(row.get('segment', '')).strip()

# # # # # # #         # ── OD payout ─────────────────────────────────────────────────────────
# # # # # # #         payin_od  = row['payin_od_rate']
# # # # # # #         old_po_od = row['payout_od_rate']

# # # # # # #         if pd.notna(payin_od) and float(payin_od) != 0:
# # # # # # #             total_od += 1
# # # # # # #             # For OD: always use COMP/SAOD formula regardless of segment column
# # # # # # #             # (OD rate only exists for COMP/SAOD rows)
# # # # # # #             new_od = compute_payout(payin_od, sp, 'COMPREHENSIVE')
# # # # # # #             new_payout_od.append(new_od)
# # # # # # #             if abs(float(old_po_od) - new_od) > 0.001:
# # # # # # #                 changed_od += 1
# # # # # # #         else:
# # # # # # #             new_payout_od.append(old_po_od)

# # # # # # #         # ── TP payout ─────────────────────────────────────────────────────────
# # # # # # #         payin_tp  = row['payin_tp_rate']
# # # # # # #         old_po_tp = row['payout_tp_rate']

# # # # # # #         if pd.notna(payin_tp) and float(payin_tp) != 0:
# # # # # # #             total_tp += 1
# # # # # # #             new_tp = compute_payout(payin_tp, sp, seg)
# # # # # # #             new_payout_tp.append(new_tp)
# # # # # # #             if abs(float(old_po_tp) - new_tp) > 0.001:
# # # # # # #                 changed_tp += 1
# # # # # # #         else:
# # # # # # #             new_payout_tp.append(old_po_tp)

# # # # # # #     df['payout_od_rate'] = new_payout_od
# # # # # # #     df['payout_tp_rate'] = new_payout_tp

# # # # # # #     # ── Save ──────────────────────────────────────────────────────────────────
# # # # # # #     df.to_excel(output_path, index=False)

# # # # # # #     print("="*65)
# # # # # # #     print(f"  DONE")
# # # # # # #     print(f"  Total rows        : {total}")
# # # # # # #     print(f"  OD rows processed : {total_od}  |  values changed: {changed_od}")
# # # # # # #     print(f"  TP rows processed : {total_tp}  |  values changed: {changed_tp}")
# # # # # # #     print(f"  Saved to          : {output_path}")
# # # # # # #     print("="*65)

# # # # # # #     # ── Quick sanity check print ──────────────────────────────────────────────
# # # # # # #     print("\n  Sample of recalculated rows:")
# # # # # # #     sample = df[df['payin_od_rate'] > 0].head(8)[[
# # # # # # #         'sub_product_name', 'segment', 'rto_group_name',
# # # # # # #         'payin_od_rate', 'payout_od_rate',
# # # # # # #         'payin_tp_rate', 'payout_tp_rate'
# # # # # # #     ]]
# # # # # # #     print(sample.to_string(index=False))
# # # # # # #     print()


# # # # # # # if __name__ == "__main__":
# # # # # # #     main()

# # # # # # # """
# # # # # # # Shriram Payin-Config — Payout Recalculator
# # # # # # # ============================================
# # # # # # # Reads SHRIRAM-Payin-Config.xlsx and recomputes payout_od_rate
# # # # # # # and payout_tp_rate using the Shriram formula:

# # # # # # #     payout = FLOOR(payin × 0.90)

# # # # # # # FLOOR drops everything after the decimal — always toward zero.
# # # # # # #   37.0  × 0.90 = 33.30  → 33
# # # # # # #   22.5  × 0.90 = 20.25  → 20
# # # # # # #   48.5  × 0.90 = 43.65  → 43
# # # # # # #   33.8  × 0.90 = 30.42  → 30

# # # # # # # Segment-specific formula overrides are available in FORMULA_OVERRIDE
# # # # # # # below if business rules change in future.

# # # # # # # Usage:
# # # # # # #     python recalculate_payout.py
# # # # # # # """

# # # # # # # import pandas as pd
# # # # # # # import math
# # # # # # # import os
# # # # # # # from datetime import datetime


# # # # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # # # #  FORMULA CONFIGURATION
# # # # # # # # ─────────────────────────────────────────────────────────────────────────────

# # # # # # # PAYOUT_RATIO = 0.90    # 90% of Payin — universal Shriram rule

# # # # # # # # Optional per-segment overrides:
# # # # # # # # key   = (sub_product_name_upper_fragment, segment_upper_fragment)
# # # # # # # # value = callable(payin) -> payout
# # # # # # # #
# # # # # # # # Example: to apply tiered TP for Two Wheeler, uncomment:
# # # # # # # # def _tiered_tp(p):
# # # # # # # #     if p <= 30:   return math.floor(p - 3)
# # # # # # # #     elif p <= 50: return math.floor(p - 4)
# # # # # # # #     else:         return math.floor(p - 5)
# # # # # # # # FORMULA_OVERRIDE = {("TWO WHEELER", "TP ONLY"): _tiered_tp}

# # # # # # # FORMULA_OVERRIDE = {}   # empty = use PAYOUT_RATIO × FLOOR for everything


# # # # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # # # #  FLOOR  (drop decimals, always toward zero — like Excel FLOOR.MATH)
# # # # # # # # ─────────────────────────────────────────────────────────────────────────────

# # # # # # # def floor_payout(value):
# # # # # # #     """
# # # # # # #     Return integer floor of value (drops all decimals, never rounds up).
# # # # # # #     Equivalent to Excel INT() or FLOOR.MATH(value, 1).

# # # # # # #     Examples
# # # # # # #     --------
# # # # # # #     floor_payout(33.80)  ->  33
# # # # # # #     floor_payout(22.50)  ->  22
# # # # # # #     floor_payout(43.65)  ->  43
# # # # # # #     floor_payout(9.00)   ->   9
# # # # # # #     floor_payout(0.315)  ->   0
# # # # # # #     """
# # # # # # #     return float(math.floor(float(value)))


# # # # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # # # #  CORE FORMULA ENGINE
# # # # # # # # ─────────────────────────────────────────────────────────────────────────────

# # # # # # # def _default(p):
# # # # # # #     """Default formula: FLOOR(payin * PAYOUT_RATIO)."""
# # # # # # #     return floor_payout(float(p) * PAYOUT_RATIO)


# # # # # # # def _pick_formula(sub_product_name, segment):
# # # # # # #     """Return the formula function for this row."""
# # # # # # #     sp  = str(sub_product_name).strip().upper()
# # # # # # #     seg = str(segment).strip().upper()
# # # # # # #     for (sp_key, seg_key), fn in FORMULA_OVERRIDE.items():
# # # # # # #         if sp_key in sp and seg_key in seg:
# # # # # # #             return fn
# # # # # # #     return _default


# # # # # # # def compute_payout(payin, sub_product_name, segment):
# # # # # # #     """
# # # # # # #     Compute payout from a single payin value using FLOOR.

# # # # # # #     Parameters
# # # # # # #     ----------
# # # # # # #     payin            : float  Payin rate (e.g. 25.0)
# # # # # # #     sub_product_name : str    e.g. 'Two Wheeler', 'Private Car'
# # # # # # #     segment          : str    e.g. 'Comprehensive', 'TP Only'

# # # # # # #     Returns
# # # # # # #     -------
# # # # # # #     float — payout rate (floored to whole number, never negative)
# # # # # # #     """
# # # # # # #     try:
# # # # # # #         p = float(payin)
# # # # # # #     except (TypeError, ValueError):
# # # # # # #         return 0.0

# # # # # # #     if p == 0:
# # # # # # #         return 0.0

# # # # # # #     fn = _pick_formula(sub_product_name, segment)
# # # # # # #     return max(0.0, fn(p))


# # # # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # # # #  MAIN
# # # # # # # # ─────────────────────────────────────────────────────────────────────────────

# # # # # # # def main():
# # # # # # #     print("\n" + "="*65)
# # # # # # #     print("  Shriram Payin-Config — Payout Recalculator")
# # # # # # #     print("  Formula: FLOOR(payin × {:.0f}%)  →  whole number only".format(
# # # # # # #           PAYOUT_RATIO * 100))
# # # # # # #     print("="*65)

# # # # # # #     input_path  = input("\nEnter path to SHRIRAM-Payin-Config.xlsx : ").strip().strip('"')
# # # # # # #     output_path = input("Enter output file path (blank = auto)   : ").strip().strip('"')

# # # # # # #     if not output_path:
# # # # # # #         base, ext = os.path.splitext(input_path)
# # # # # # #         ts = datetime.now().strftime("%Y%m%d_%H%M%S")
# # # # # # #         output_path = f"{base}_recalculated_{ts}{ext}"

# # # # # # #     print(f"\n  Reading  : {input_path}")
# # # # # # #     df = pd.read_excel(input_path)
# # # # # # #     df.columns = [c.strip() for c in df.columns]
# # # # # # #     total = len(df)
# # # # # # #     print(f"  Rows     : {total}")

# # # # # # #     # Validate required columns
# # # # # # #     required = ['payin_od_rate', 'payin_tp_rate',
# # # # # # #                 'payout_od_rate', 'payout_tp_rate',
# # # # # # #                 'sub_product_name', 'segment']
# # # # # # #     missing = [c for c in required if c not in df.columns]
# # # # # # #     if missing:
# # # # # # #         print(f"\n[ERROR] Missing required columns: {missing}")
# # # # # # #         return

# # # # # # #     if FORMULA_OVERRIDE:
# # # # # # #         print(f"  Overrides active: {list(FORMULA_OVERRIDE.keys())}")

# # # # # # #     print(f"\n  Processing rows ...\n")

# # # # # # #     changed_od   = 0
# # # # # # #     changed_tp   = 0
# # # # # # #     processed_od = 0
# # # # # # #     processed_tp = 0
# # # # # # #     new_payout_od = []
# # # # # # #     new_payout_tp = []

# # # # # # #     for _, row in df.iterrows():
# # # # # # #         sp  = str(row.get('sub_product_name', '')).strip()
# # # # # # #         seg = str(row.get('segment', '')).strip()

# # # # # # #         # ── OD / COMP column ──────────────────────────────────────────────────
# # # # # # #         payin_od = row['payin_od_rate']
# # # # # # #         old_od   = row['payout_od_rate']

# # # # # # #         if pd.notna(payin_od) and float(payin_od) != 0:
# # # # # # #             processed_od += 1
# # # # # # #             new_od = compute_payout(payin_od, sp, seg)
# # # # # # #             new_payout_od.append(new_od)
# # # # # # #             if abs(float(old_od) - new_od) > 0.001:
# # # # # # #                 changed_od += 1
# # # # # # #         else:
# # # # # # #             new_payout_od.append(0.0 if pd.isna(old_od) else old_od)

# # # # # # #         # ── TP column ─────────────────────────────────────────────────────────
# # # # # # #         payin_tp = row['payin_tp_rate']
# # # # # # #         old_tp   = row['payout_tp_rate']

# # # # # # #         if pd.notna(payin_tp) and float(payin_tp) != 0:
# # # # # # #             processed_tp += 1
# # # # # # #             new_tp = compute_payout(payin_tp, sp, seg)
# # # # # # #             new_payout_tp.append(new_tp)
# # # # # # #             if abs(float(old_tp) - new_tp) > 0.001:
# # # # # # #                 changed_tp += 1
# # # # # # #         else:
# # # # # # #             new_payout_tp.append(0.0 if pd.isna(old_tp) else old_tp)

# # # # # # #     df['payout_od_rate'] = new_payout_od
# # # # # # #     df['payout_tp_rate'] = new_payout_tp

# # # # # # #     # Save
# # # # # # #     df.to_excel(output_path, index=False)

# # # # # # #     print(f"{'='*65}")
# # # # # # #     print(f"  COMPLETED")
# # # # # # #     print(f"  Total rows           : {total}")
# # # # # # #     print(f"  OD rows recalculated : {processed_od}   values changed: {changed_od}")
# # # # # # #     print(f"  TP rows recalculated : {processed_tp}   values changed: {changed_tp}")
# # # # # # #     print(f"  Output saved to      : {output_path}")
# # # # # # #     print(f"{'='*65}")

# # # # # # #     # Sanity check: show sample rows
# # # # # # #     print("\n  Sample output (first 10 non-zero OD rows):\n")
# # # # # # #     sample = df[df['payin_od_rate'] > 0].head(10)[[
# # # # # # #         'sub_product_name', 'segment', 'rto_group_name',
# # # # # # #         'payin_od_rate', 'payout_od_rate',
# # # # # # #         'payin_tp_rate', 'payout_tp_rate',
# # # # # # #     ]]
# # # # # # #     print(f"  {'Sub Product':<22} {'Segment':<16} {'Location':<20} "
# # # # # # #           f"{'Pay_OD':>7} {'PO_OD':>7} {'Pay_TP':>7} {'PO_TP':>7}")
# # # # # # #     print(f"  {'-'*22} {'-'*16} {'-'*20} {'-'*7} {'-'*7} {'-'*7} {'-'*7}")
# # # # # # #     for _, r in sample.iterrows():
# # # # # # #         print(f"  {str(r['sub_product_name']):<22} "
# # # # # # #               f"{str(r['segment']):<16} "
# # # # # # #               f"{str(r['rto_group_name']):<20} "
# # # # # # #               f"{r['payin_od_rate']:>7.2f} "
# # # # # # #               f"{r['payout_od_rate']:>7.0f} "
# # # # # # #               f"{r['payin_tp_rate']:>7.2f} "
# # # # # # #               f"{r['payout_tp_rate']:>7.0f}")

# # # # # # #     # FLOOR illustration table
# # # # # # #     print("\n\n  FLOOR illustration (payin × 90% → floor to whole number):\n")
# # # # # # #     print(f"  {'Payin':>8}  {'× 0.90':>10}  {'FLOOR':>8}")
# # # # # # #     print(f"  {'-'*8}  {'-'*10}  {'-'*8}")
# # # # # # #     for p in [37, 25, 22.5, 48.5, 53.5, 33.8, 12.5, 10, 7.5, 0.35, 0.55]:
# # # # # # #         raw = p * 0.90
# # # # # # #         print(f"  {p:>8}  {raw:>10.4f}  {math.floor(raw):>8}")
# # # # # # #     print()


# # # # # # # if __name__ == "__main__":
# # # # # # #     main()

# # # # # #     """
# # # # # #     Payin-Config — Payout Recalculator for All Companies
# # # # # #     ====================================================
# # # # # #     Reads Payin-Config.xlsx and recomputes payout_od_rate and payout_tp_rate
# # # # # #     using rules from payout_rules.json.

# # # # # #     Rules are matched based on LOB (sub_product_name), SEGMENT (derived from segment and OD/TP),
# # # # # #     INSURER (insurer column), and LOCATION (rto_group_name contains 'ODISHA' for overrides).
# # # # # #     FLOOR drops everything after the decimal — always toward zero.

# # # # # #     Usage:
# # # # # #         python recalculate_payout.py
# # # # # #     """

# # # # # #     import pandas as pd
# # # # # #     import math
# # # # # #     import os
# # # # # #     import json
# # # # # #     import re
# # # # # #     from datetime import datetime


# # # # # #     # ─────────────────────────────────────────────────────────────────────────────
# # # # # #     #  LOB MAPPING (Excel sub_product_name -> JSON LOB)
# # # # # #     # ─────────────────────────────────────────────────────────────────────────────

# # # # # #     LOB_MAP = {
# # # # # #         "TWO WHEELER": "TW",
# # # # # #         "PRIVATE CAR": "PVT CAR",
# # # # # #         "GCV, PCV 3W": "GCV, PCV 3W",
# # # # # #         "BUS": "BUS",
# # # # # #         "TAXI": "TAXI",
# # # # # #         "MISD": "MISD",  # or "Misd, Tractor"
# # # # # #         # Add more mappings as needed, e.g., {"Misd, Tractor": "MISD"}
# # # # # #     }


# # # # # #     # ─────────────────────────────────────────────────────────────────────────────
# # # # # #     #  FLOOR  (drop decimals, always toward zero — like Excel FLOOR.MATH)
# # # # # #     # ─────────────────────────────────────────────────────────────────────────────

# # # # # #     def floor_payout(value):
# # # # # #         """
# # # # # #         Return integer floor of value (drops all decimals, never rounds up).
# # # # # #         Equivalent to Excel INT() or FLOOR.MATH(value, 1).

# # # # # #         Examples
# # # # # #         --------
# # # # # #         floor_payout(33.80)  ->  33
# # # # # #         floor_payout(22.50)  ->  22
# # # # # #         floor_payout(43.65)  ->  43
# # # # # #         floor_payout(9.00)   ->   9
# # # # # #         floor_payout(0.315)  ->   0
# # # # # #         """
# # # # # #         return float(math.floor(float(value)))


# # # # # #     # ─────────────────────────────────────────────────────────────────────────────
# # # # # #     #  PO STRING PARSER
# # # # # #     # ─────────────────────────────────────────────────────────────────────────────

# # # # # #     def parse_po_to_payout(po_str, p):
# # # # # #         """
# # # # # #         Parse PO string to payout value (floored).
# # # # # #         """
# # # # # #         po_str = str(po_str).strip().upper()
# # # # # #         if "% OF PAYIN" in po_str:
# # # # # #             percent = float(re.search(r'(\d+)%', po_str).group(1))
# # # # # #             return floor_payout(p * (percent / 100))
# # # # # #         elif "PAYIN + 1" in po_str:
# # # # # #             return floor_payout(p + 1)
# # # # # #         elif "LESS 2% OF PAYIN" in po_str:
# # # # # #             return floor_payout(p * 0.98)
# # # # # #         elif po_str.startswith("-") and po_str.endswith("%"):
# # # # # #             ded = abs(float(re.search(r'-(\d+)%', po_str).group(1)))
# # # # # #             return floor_payout(p - ded)
# # # # # #         elif "21% PO" in po_str:
# # # # # #             return floor_payout(p * 0.21)
# # # # # #         else:
# # # # # #             # Fallback to 90%
# # # # # #             return floor_payout(p * 0.90)


# # # # # #     # ─────────────────────────────────────────────────────────────────────────────
# # # # # #     #  RULE SELECTION HELPER
# # # # # #     # ─────────────────────────────────────────────────────────────────────────────

# # # # # #     def select_po(rules_list, p):
# # # # # #         """
# # # # # #         Select the PO from candidate rules based on payin tiers in REMARKS.
# # # # # #         """
# # # # # #         if not rules_list:
# # # # # #             return None
# # # # # #         for r in rules_list:
# # # # # #             rem = str(r.get("REMARKS", "")).upper()
# # # # # #             if not rem.strip() or "NIL" in rem:
# # # # # #                 return r["PO"]
# # # # # #             # Parse "Payin Below XX%"
# # # # # #             match_below = re.search(r'BELOW (\d+)%', rem)
# # # # # #             if match_below and p <= float(match_below.group(1)):
# # # # # #                 return r["PO"]
# # # # # #             # Parse "Payin XX% to YY%"
# # # # # #             match_range = re.search(r'(\d+)% TO (\d+)%', rem)
# # # # # #             if match_range:
# # # # # #                 low = float(match_range.group(1))
# # # # # #                 high = float(match_range.group(2))
# # # # # #                 if low <= p <= high:
# # # # # #                     return r["PO"]
# # # # # #             # Parse "Payin Above XX%"
# # # # # #             match_above = re.search(r'ABOVE (\d+)%', rem)
# # # # # #             if match_above and p > float(match_above.group(1)):
# # # # # #                 return r["PO"]
# # # # # #         # Fallback to first
# # # # # #         return rules_list[0]["PO"]


# # # # # #     # ─────────────────────────────────────────────────────────────────────────────
# # # # # #     #  CORE FORMULA ENGINE
# # # # # #     # ─────────────────────────────────────────────────────────────────────────────

# # # # # #     def compute_payout(payin, sub_product_name, segment, insurer, location, rules, is_od=True):
# # # # # #         """
# # # # # #         Compute payout from a single payin value using JSON rules (floored).

# # # # # #         Parameters
# # # # # #         ----------
# # # # # #         payin            : float  Payin rate (e.g. 25.0)
# # # # # #         sub_product_name : str    e.g. 'Two Wheeler', 'Private Car'
# # # # # #         segment          : str    e.g. 'Comprehensive', 'TP Only'
# # # # # #         insurer          : str    e.g. 'TATA', 'ZUNO'
# # # # # #         location         : str    e.g. rto_group_name
# # # # # #         rules            : list   JSON rules list
# # # # # #         is_od            : bool   True for OD/COMP, False for TP

# # # # # #         Returns
# # # # # #         -------
# # # # # #         float — payout rate (floored to whole number, never negative)
# # # # # #         """
# # # # # #         try:
# # # # # #             p = float(payin)
# # # # # #         except (TypeError, ValueError):
# # # # # #             return 0.0

# # # # # #         if p == 0:
# # # # # #             return 0.0

# # # # # #         sp = str(sub_product_name).strip().upper()
# # # # # #         seg = str(segment).strip().upper()
# # # # # #         ins = str(insurer).strip().upper() if pd.notna(insurer) else ""
# # # # # #         loc = str(location).strip().upper() if pd.notna(location) else ""

# # # # # #         # Map to LOB
# # # # # #         lob_key = None
# # # # # #         for full_name, short in LOB_MAP.items():
# # # # # #             if full_name in sp:
# # # # # #                 lob_key = short
# # # # # #                 break
# # # # # #         if not lob_key:
# # # # # #             # Default fallback
# # # # # #             return floor_payout(p * 0.90)

# # # # # #         # Derive JSON SEGMENT
# # # # # #         if is_od:
# # # # # #             if "COMP" in seg or "OD" in seg or "SAOD" in seg:
# # # # # #                 if lob_key == "TW":
# # # # # #                     json_seg = "TW SAOD + COMP"
# # # # # #                 elif lob_key == "PVT CAR":
# # # # # #                     json_seg = "PVT CAR COMP + SAOD"
# # # # # #                 else:
# # # # # #                     json_seg = seg
# # # # # #             else:
# # # # # #                 json_seg = seg
# # # # # #         else:
# # # # # #             if "TP" in seg:
# # # # # #                 if lob_key == "TW":
# # # # # #                     json_seg = "TW TP"
# # # # # #                 elif lob_key == "PVT CAR":
# # # # # #                     json_seg = "PVT CAR TP"
# # # # # #                 else:
# # # # # #                     json_seg = seg
# # # # # #             else:
# # # # # #                 json_seg = seg

# # # # # #         # Find matching rules by LOB and SEGMENT
# # # # # #         matching_rules = [r for r in rules if r.get("LOB") == lob_key and r.get("SEGMENT") == json_seg]

# # # # # #         # Filter candidates by INSURER (handle specifics, then rest)
# # # # # #         specific_rules = []
# # # # # #         for r in matching_rules:
# # # # # #             if r.get("INSURER") == "Rest of Companies":
# # # # # #                 continue
# # # # # #             ins_rule = str(r.get("INSURER", "")).upper()
# # # # # #             if ins_rule == "ALL COMPANIES" or ins in [i.strip().upper() for i in ins_rule.split(",")]:
# # # # # #                 specific_rules.append(r)

# # # # # #         if specific_rules:
# # # # # #             candidate_rules = specific_rules
# # # # # #         else:
# # # # # #             rest_rules = [r for r in matching_rules if "rest of companies" in str(r.get("INSURER", "")).lower()]
# # # # # #             if rest_rules:
# # # # # #                 candidate_rules = rest_rules
# # # # # #             else:
# # # # # #                 candidate_rules = [r for r in matching_rules if r.get("INSURER") == "All Companies"]

# # # # # #         # Select PO
# # # # # #         selected_po = select_po(candidate_rules, p)
# # # # # #         if selected_po is None:
# # # # # #             # No rule, default
# # # # # #             p_out = floor_payout(p * 0.90)
# # # # # #         else:
# # # # # #             p_out = parse_po_to_payout(selected_po, p)

# # # # # #         # Odisha override (additional deduction)
# # # # # #         if "ODISHA" in loc:
# # # # # #             odisha_rules = [r for r in rules if r.get("LOCATION") == "ODISHA" and
# # # # # #                             r.get("SEGMENT") == "ALL SEGMENT" and
# # # # # #                             r.get("INSURER") == "All Companies"]
# # # # # #             selected_ded_po = select_po(odisha_rules, p)
# # # # # #             if selected_ded_po:
# # # # # #                 ded_str = str(selected_ded_po).upper()
# # # # # #                 ded_match = re.search(r'-(\d+)%', ded_str)
# # # # # #                 if ded_match:
# # # # # #                     ded = float(ded_match.group(1))
# # # # # #                     p_out = floor_payout(p_out - ded)

# # # # # #         return max(0.0, p_out)


# # # # # #     # ─────────────────────────────────────────────────────────────────────────────
# # # # # #     #  MAIN
# # # # # #     # ─────────────────────────────────────────────────────────────────────────────

# # # # # #     def main():
# # # # # #         print("\n" + "="*70)
# # # # # #         print("  Payin-Config — Payout Recalculator for All Companies")
# # # # # #         print("  Using rules from JSON (multipliers, deductions, tiers)")
# # # # # #         print("="*70)

# # # # # #         json_path = input("\nEnter path to payout_rules.json : ").strip().strip('"')
# # # # # #         try:
# # # # # #             with open(json_path, 'r') as f:
# # # # # #                 rules = json.load(f)
# # # # # #             print(f"  Loaded {len(rules)} rules from {json_path}")
# # # # # #         except Exception as e:
# # # # # #             print(f"\n[ERROR] Failed to load JSON: {e}")
# # # # # #             return

# # # # # #         input_path = input("Enter path to Payin-Config.xlsx : ").strip().strip('"')
# # # # # #         output_path = input("Enter output file path (blank = auto)   : ").strip().strip('"')

# # # # # #         if not output_path:
# # # # # #             base, ext = os.path.splitext(input_path)
# # # # # #             ts = datetime.now().strftime("%Y%m%d_%H%M%S")
# # # # # #             output_path = f"{base}_recalculated_{ts}{ext}"

# # # # # #         print(f"\n  Reading  : {input_path}")
# # # # # #         df = pd.read_excel(input_path)
# # # # # #         df.columns = [c.strip() for c in df.columns]
# # # # # #         total = len(df)
# # # # # #         print(f"  Rows     : {total}")

# # # # # #         # Validate required columns
# # # # # #         required = ['payin_od_rate', 'payin_tp_rate',
# # # # # #                     'payout_od_rate', 'payout_tp_rate',
# # # # # #                     'sub_product_name', 'segment', 'insurer', 'rto_group_name']
# # # # # #         missing = [c for c in required if c not in df.columns]
# # # # # #         if missing:
# # # # # #             print(f"\n[ERROR] Missing required columns: {missing}")
# # # # # #             print("  Ensure 'insurer' and 'rto_group_name' are present.")
# # # # # #             return

# # # # # #         print(f"\n  Processing rows using JSON rules ...\n")

# # # # # #         changed_od = 0
# # # # # #         changed_tp = 0
# # # # # #         processed_od = 0
# # # # # #         processed_tp = 0
# # # # # #         new_payout_od = []
# # # # # #         new_payout_tp = []

# # # # # #         for _, row in df.iterrows():
# # # # # #             sp = row['sub_product_name']
# # # # # #             seg = row['segment']
# # # # # #             ins = row['insurer']
# # # # # #             loc = row['rto_group_name']

# # # # # #             # ── OD / COMP column ──────────────────────────────────────────────────
# # # # # #             payin_od = row['payin_od_rate']
# # # # # #             old_od = row['payout_od_rate']

# # # # # #             if pd.notna(payin_od) and float(payin_od) != 0:
# # # # # #                 processed_od += 1
# # # # # #                 new_od = compute_payout(payin_od, sp, seg, ins, loc, rules, is_od=True)
# # # # # #                 new_payout_od.append(new_od)
# # # # # #                 if abs(float(old_od) - new_od) > 0.001:
# # # # # #                     changed_od += 1
# # # # # #             else:
# # # # # #                 new_payout_od.append(0.0 if pd.isna(old_od) else old_od)

# # # # # #             # ── TP column ─────────────────────────────────────────────────────────
# # # # # #             payin_tp = row['payin_tp_rate']
# # # # # #             old_tp = row['payout_tp_rate']

# # # # # #             if pd.notna(payin_tp) and float(payin_tp) != 0:
# # # # # #                 processed_tp += 1
# # # # # #                 new_tp = compute_payout(payin_tp, sp, seg, ins, loc, rules, is_od=False)
# # # # # #                 new_payout_tp.append(new_tp)
# # # # # #                 if abs(float(old_tp) - new_tp) > 0.001:
# # # # # #                     changed_tp += 1
# # # # # #             else:
# # # # # #                 new_payout_tp.append(0.0 if pd.isna(old_tp) else old_tp)

# # # # # #         df['payout_od_rate'] = new_payout_od
# # # # # #         df['payout_tp_rate'] = new_payout_tp

# # # # # #         # Save
# # # # # #         df.to_excel(output_path, index=False)

# # # # # #         print(f"{'='*70}")
# # # # # #         print(f"  COMPLETED")
# # # # # #         print(f"  Total rows           : {total}")
# # # # # #         print(f"  OD rows recalculated : {processed_od}   values changed: {changed_od}")
# # # # # #         print(f"  TP rows recalculated : {processed_tp}   values changed: {changed_tp}")
# # # # # #         print(f"  Output saved to      : {output_path}")
# # # # # #         print(f"{'='*70}")

# # # # # #         # Sanity check: show sample rows
# # # # # #         print("\n  Sample output (first 10 non-zero OD rows):\n")
# # # # # #         sample_cols = ['sub_product_name', 'segment', 'insurer', 'rto_group_name',
# # # # # #                     'payin_od_rate', 'payout_od_rate',
# # # # # #                     'payin_tp_rate', 'payout_tp_rate']
# # # # # #         sample = df[df['payin_od_rate'] > 0].head(10)[sample_cols]
# # # # # #         print(f"  {'Sub Product':<15} {'Segment':<12} {'Insurer':<10} {'Location':<15} "
# # # # # #             f"{'Pay_OD':>7} {'PO_OD':>7} {'Pay_TP':>7} {'PO_TP':>7}")
# # # # # #         print(f"  {'-'*15} {'-'*12} {'-'*10} {'-'*15} {'-'*7} {'-'*7} {'-'*7} {'-'*7}")
# # # # # #         for _, r in sample.iterrows():
# # # # # #             print(f"  {str(r['sub_product_name']):<15} "
# # # # # #                 f"{str(r['segment']):<12} "
# # # # # #                 f"{str(r['insurer']):<10} "
# # # # # #                 f"{str(r['rto_group_name']):<15} "
# # # # # #                 f"{r['payin_od_rate']:>7.2f} "
# # # # # #                 f"{r['payout_od_rate']:>7.0f} "
# # # # # #                 f"{r['payin_tp_rate']:>7.2f} "
# # # # # #                 f"{r['payout_tp_rate']:>7.0f}")

# # # # # #         # FLOOR illustration table (example)
# # # # # #         print("\n\n  FLOOR illustration (examples: payin × mult/ded → floor to whole number):\n")
# # # # # #         print(f"  {'Payin':>8}  {'Operation':>15}  {'FLOOR':>8}")
# # # # # #         print(f"  {'-'*8}  {'-'*15}  {'-'*8}")
# # # # # #         examples = [
# # # # # #             (37, "× 0.88", 37 * 0.88),
# # # # # #             (25, "- 3", 25 - 3),
# # # # # #             (22.5, "× 0.90", 22.5 * 0.90),
# # # # # #             (48.5, "- 5", 48.5 - 5),
# # # # # #             (53.5, "+ 1", 53.5 + 1),
# # # # # #             (33.8, "× 0.85", 33.8 * 0.85),
# # # # # #             (12.5, "- 1", 12.5 - 1),  # e.g., Odisha
# # # # # #             (10, "× 0.98", 10 * 0.98),
# # # # # #             (7.5, "× 0.21", 7.5 * 0.21),
# # # # # #             (0.35, "default ×0.90", 0.35 * 0.90),
# # # # # #         ]
# # # # # #         for p, op, raw in examples:
# # # # # #             floored = floor_payout(raw)
# # # # # #             print(f"  {p:>8}  {op:>15}  {floored:>8}")
# # # # # #         print()


# # # # # #     if __name__ == "__main__":
# # # # # #         main()

# # # # # """
# # # # # Payin-Config — Payout Recalculator
# # # # # ===================================
# # # # # Reads PayinConfig.xlsx + master Vehicle Type table (hardcoded from provided data)
# # # # # and recomputes payout_od_rate / payout_tp_rate using JSON payout rules.

# # # # # Usage:
# # # # #     python recalculate_payout.py
# # # # # """

# # # # # import pandas as pd
# # # # # import math
# # # # # import os
# # # # # import json
# # # # # import re
# # # # # from datetime import datetime

# # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # #  MASTER — Vehicle Type lookup (from master file Vehicle type worksheet)
# # # # # # ─────────────────────────────────────────────────────────────────────────────

# # # # # VEHICLE_TYPE_MASTER = {
# # # # #     1:  {"vehicle_type": "Agriculture Tractor",          "sub_product_name": "Miscellaneous Vehicle"},
# # # # #     2:  {"vehicle_type": "Non Tractor",                  "sub_product_name": "Miscellaneous Vehicle"},
# # # # #     3:  {"vehicle_type": "Truck",                        "sub_product_name": "Goods Vehicle"},
# # # # #     4:  {"vehicle_type": "Good Carring Tractor",         "sub_product_name": "Goods Vehicle"},
# # # # #     5:  {"vehicle_type": "Tanker",                       "sub_product_name": "Goods Vehicle"},
# # # # #     6:  {"vehicle_type": "Pickup",                       "sub_product_name": "Goods Vehicle"},
# # # # #     7:  {"vehicle_type": "GCV 3W Delivery Van",          "sub_product_name": "Goods Vehicle"},
# # # # #     8:  {"vehicle_type": "Taxi_CAB",                     "sub_product_name": "Passenger Vehicle"},
# # # # #     9:  {"vehicle_type": "Electric Rikshaw",             "sub_product_name": "Passenger Vehicle"},
# # # # #     10: {"vehicle_type": "Tempo Traveller",              "sub_product_name": "Passenger Vehicle"},
# # # # #     11: {"vehicle_type": "School Bus",                   "sub_product_name": "Passenger Vehicle"},
# # # # #     12: {"vehicle_type": "Passanger Bus",                "sub_product_name": "Passenger Vehicle"},
# # # # #     13: {"vehicle_type": "Auto rikshaw",                 "sub_product_name": "Passenger Vehicle"},
# # # # #     14: {"vehicle_type": "3W Tipper",                    "sub_product_name": "Goods Vehicle"},
# # # # #     15: {"vehicle_type": "PCV 2W",                       "sub_product_name": "Passenger Vehicle"},
# # # # #     16: {"vehicle_type": "GCV 2W",                       "sub_product_name": "Goods Vehicle"},
# # # # #     17: {"vehicle_type": "TW Scooter",                   "sub_product_name": "Two Wheeler"},
# # # # #     18: {"vehicle_type": "TW Bike",                      "sub_product_name": "Two Wheeler"},
# # # # #     19: {"vehicle_type": "Private Car",                  "sub_product_name": "Private Car"},
# # # # #     20: {"vehicle_type": "TW Electric Bike",             "sub_product_name": "Two Wheeler"},
# # # # #     21: {"vehicle_type": "Electric GCV 3W Delivery Van", "sub_product_name": "Goods Vehicle"},
# # # # #     22: {"vehicle_type": "Private Car Electric",         "sub_product_name": "Private Car"},
# # # # #     23: {"vehicle_type": "Trailer",                      "sub_product_name": "Goods Vehicle"},
# # # # #     24: {"vehicle_type": "Electric Pickup",              "sub_product_name": "Goods Vehicle"},
# # # # #     25: {"vehicle_type": "Tipper",                       "sub_product_name": "Goods Vehicle"},
# # # # #     26: {"vehicle_type": "TW Electric Scooter",          "sub_product_name": "Two Wheeler"},
# # # # #     27: {"vehicle_type": "PC Petrol / Electric Hybrid",  "sub_product_name": "Private Car"},
# # # # #     28: {"vehicle_type": "Staff Bus",                    "sub_product_name": "Passenger Vehicle"},
# # # # #     29: {"vehicle_type": "Agriculture Harvester",        "sub_product_name": "Miscellaneous Vehicle"},
# # # # #     30: {"vehicle_type": "Electric PCV 2W",              "sub_product_name": "Passenger Vehicle"},
# # # # #     31: {"vehicle_type": "Electric Taxi_CAB",            "sub_product_name": "Passenger Vehicle"},
# # # # #     32: {"vehicle_type": "Route Bus",                    "sub_product_name": "Passenger Vehicle"},
# # # # # }

# # # # # # IDs considered as TAXI
# # # # # TAXI_VEHICLE_IDS = {8, 31}  # Taxi_CAB, Electric Taxi_CAB

# # # # # # IDs considered as STAFF BUS
# # # # # STAFF_BUS_IDS = {28}  # Staff Bus

# # # # # # IDs considered as SCHOOL BUS
# # # # # SCHOOL_BUS_IDS = {11}  # School Bus

# # # # # # IDs considered as BUS (any bus — route/passenger)
# # # # # ROUTE_BUS_IDS = {12, 32}  # Passanger Bus, Route Bus

# # # # # # IDs considered as GCV 3-Wheeler goods
# # # # # GCV_3W_IDS = {7, 14, 21}  # GCV 3W Delivery Van, 3W Tipper, Electric GCV 3W

# # # # # # IDs considered as Passenger 3-Wheeler (auto etc.)
# # # # # PCV_3W_IDS = {9, 13, 15, 30}  # Electric Rikshaw, Auto rikshaw, PCV 2W, Electric PCV 2W

# # # # # # Insurers to match for Upto 2.5 GVW special rule
# # # # # SPECIAL_GCV_INSURERS = {"RELIANCE", "SBI"}

# # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # #  FLOOR
# # # # # # ─────────────────────────────────────────────────────────────────────────────

# # # # # def floor_payout(value):
# # # # #     return float(math.floor(float(value)))


# # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # #  PO STRING PARSER
# # # # # # ─────────────────────────────────────────────────────────────────────────────

# # # # # def parse_po_to_payout(po_str, p):
# # # # #     po_str = str(po_str).strip().upper()

# # # # #     if re.search(r'\d+%\s*OF PAYIN', po_str):
# # # # #         percent = float(re.search(r'(\d+(?:\.\d+)?)%', po_str).group(1))
# # # # #         return floor_payout(p * (percent / 100))

# # # # #     if "PAYIN + 1" in po_str:
# # # # #         return floor_payout(p + 1)

# # # # #     if "LESS 2% OF PAYIN" in po_str:
# # # # #         return floor_payout(p - 2)

# # # # #     # "-3%", "-4%", "-5%"
# # # # #     m = re.fullmatch(r'-(\d+(?:\.\d+)?)%', po_str)
# # # # #     if m:
# # # # #         ded = float(m.group(1))
# # # # #         return floor_payout(p - ded)

# # # # #     # "21% PO" — fixed payout
# # # # #     m = re.search(r'(\d+(?:\.\d+)?)%\s*PO', po_str)
# # # # #     if m:
# # # # #         return floor_payout(float(m.group(1)))

# # # # #     # Fallback
# # # # #     return floor_payout(p * 0.90)


# # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # #  RULE SELECTION — slab-based matching on REMARKS
# # # # # # ─────────────────────────────────────────────────────────────────────────────

# # # # # def select_po(rules_list, p):
# # # # #     if not rules_list:
# # # # #         return None
# # # # #     for r in rules_list:
# # # # #         rem = str(r.get("REMARKS", "")).upper().strip()
# # # # #         if not rem or rem == "NIL" or rem == "ALL FUEL":
# # # # #             return r["PO"]
# # # # #         m_below = re.search(r'BELOW\s+(\d+)%', rem)
# # # # #         if m_below and p <= float(m_below.group(1)):
# # # # #             return r["PO"]
# # # # #         m_range = re.search(r'(\d+)%\s+TO\s+(\d+)%', rem)
# # # # #         if m_range:
# # # # #             lo, hi = float(m_range.group(1)), float(m_range.group(2))
# # # # #             if lo <= p <= hi:
# # # # #                 return r["PO"]
# # # # #         m_above = re.search(r'ABOVE\s+(\d+)%', rem)
# # # # #         if m_above and p > float(m_above.group(1)):
# # # # #             return r["PO"]
# # # # #     return rules_list[0]["PO"]


# # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # #  INSURER MATCHING HELPER
# # # # # # ─────────────────────────────────────────────────────────────────────────────

# # # # # def insurer_matches(rule_insurer_str, ins):
# # # # #     """Check if the row's insurer matches a rule's INSURER field."""
# # # # #     ri = str(rule_insurer_str).strip().upper()
# # # # #     if ri == "ALL COMPANIES":
# # # # #         return True
# # # # #     rule_insurers = [x.strip().upper() for x in ri.split(",")]
# # # # #     return ins.upper() in rule_insurers


# # # # # def filter_by_insurer(rules, ins):
# # # # #     """
# # # # #     Return best matching rules for given insurer:
# # # # #     1. Specific match (not 'All Companies', not 'Rest of Companies')
# # # # #     2. 'All Companies'
# # # # #     3. 'Rest of Companies'
# # # # #     """
# # # # #     specific = [r for r in rules
# # # # #                 if str(r.get("INSURER","")).strip().upper() not in ("ALL COMPANIES","REST OF COMPANIES")
# # # # #                 and insurer_matches(r.get("INSURER",""), ins)]
# # # # #     if specific:
# # # # #         return specific

# # # # #     all_co = [r for r in rules if str(r.get("INSURER","")).strip().upper() == "ALL COMPANIES"]
# # # # #     if all_co:
# # # # #         return all_co

# # # # #     rest = [r for r in rules if str(r.get("INSURER","")).strip().upper() == "REST OF COMPANIES"]
# # # # #     return rest


# # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # #  DETERMINE JSON SEGMENT from row data
# # # # # # ─────────────────────────────────────────────────────────────────────────────

# # # # # def get_json_lob_and_segment(sub_product_name, segment, vehicle_type_id,
# # # # #                               from_wt, to_wt, company_code, is_od):
# # # # #     """
# # # # #     Returns (lob, json_segment) tuple for rule lookup.
# # # # #     lob is the JSON LOB string.
# # # # #     json_segment is the JSON SEGMENT string.
# # # # #     Returns (None, None) if not mappable.
# # # # #     """
# # # # #     sp = str(sub_product_name).strip()
# # # # #     seg = str(segment).strip().upper()
# # # # #     vt_id = int(vehicle_type_id) if pd.notna(vehicle_type_id) else 0
# # # # #     ins_upper = str(company_code).strip().upper()

# # # # #     # ── TWO WHEELER ──────────────────────────────────────────────────────────
# # # # #     if sp == "Two Wheeler":
# # # # #         lob = "TW"
# # # # #         if is_od or "COMP" in seg or "SAOD" in seg or "OD" in seg:
# # # # #             return lob, "TW SAOD + COMP"
# # # # #         else:  # TP Only
# # # # #             return lob, "TW TP"

# # # # #     # ── PRIVATE CAR ──────────────────────────────────────────────────────────
# # # # #     if sp == "Private Car":
# # # # #         lob = "PVT CAR"
# # # # #         if is_od or "COMP" in seg or "SAOD" in seg or "OD" in seg:
# # # # #             return lob, "PVT CAR COMP + SAOD"
# # # # #         else:
# # # # #             return lob, "PVT CAR TP"

# # # # #     # ── PASSENGER VEHICLE ────────────────────────────────────────────────────
# # # # #     if sp == "Passenger Vehicle":
# # # # #         # TAXI
# # # # #         if vt_id in TAXI_VEHICLE_IDS:
# # # # #             return "TAXI", "TAXI"

# # # # #         # STAFF BUS
# # # # #         if vt_id in STAFF_BUS_IDS:
# # # # #             return "BUS", "STAFF BUS"

# # # # #         # SCHOOL BUS
# # # # #         if vt_id in SCHOOL_BUS_IDS:
# # # # #             return "BUS", "SCHOOL BUS"

# # # # #         # ROUTE/PASSENGER BUS → treat as STAFF BUS (general BUS)
# # # # #         if vt_id in ROUTE_BUS_IDS:
# # # # #             return "BUS", "STAFF BUS"

# # # # #         # 3-wheelers (Auto, Electric Rikshaw, PCV 2W etc.)
# # # # #         if vt_id in PCV_3W_IDS:
# # # # #             return "GCV, PCV 3W", "All GVW & PCV 3W, GCV 3W"

# # # # #         # Tempo Traveller — treat as staff bus
# # # # #         return "BUS", "STAFF BUS"

# # # # #     # ── GOODS VEHICLE ────────────────────────────────────────────────────────
# # # # #     if sp == "Goods Vehicle":
# # # # #         from_w = float(from_wt) if pd.notna(from_wt) else 0
# # # # #         to_w   = float(to_wt)   if pd.notna(to_wt)   else 99999

# # # # #         # 3-Wheeler goods
# # # # #         if vt_id in GCV_3W_IDS:
# # # # #             return "GCV, PCV 3W", "All GVW & PCV 3W, GCV 3W"

# # # # #         # Upto 2.5T GVW + special insurers
# # # # #         if from_w == 0 and to_w <= 2500 and ins_upper in SPECIAL_GCV_INSURERS:
# # # # #             return "GCV, PCV 3W", "Upto 2.5 GVW"

# # # # #         # Everything else (inc. upto 2.5T with other insurers)
# # # # #         return "GCV, PCV 3W", "All GVW & PCV 3W, GCV 3W"

# # # # #     # ── MISCELLANEOUS VEHICLE (Tractor, Harvester etc.) ───────────────────────
# # # # #     if sp == "Miscellaneous Vehicle":
# # # # #         return "MISD", "Misd, Tractor"

# # # # #     return None, None


# # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # #  CORE FORMULA ENGINE
# # # # # # ─────────────────────────────────────────────────────────────────────────────

# # # # # def compute_payout(payin, sub_product_name, segment, company_code,
# # # # #                    rto_group_name, vehicle_type_id, from_wt, to_wt,
# # # # #                    rules, is_od=True):
# # # # #     try:
# # # # #         p = float(payin)
# # # # #     except (TypeError, ValueError):
# # # # #         return 0.0
# # # # #     if p == 0:
# # # # #         return 0.0

# # # # #     ins = str(company_code).strip().upper() if pd.notna(company_code) else ""
# # # # #     loc = str(rto_group_name).strip().upper() if pd.notna(rto_group_name) else ""

# # # # #     lob, json_seg = get_json_lob_and_segment(
# # # # #         sub_product_name, segment, vehicle_type_id, from_wt, to_wt, company_code, is_od
# # # # #     )

# # # # #     if lob is None:
# # # # #         p_out = floor_payout(p * 0.90)
# # # # #     else:
# # # # #         # Find rules by LOB + SEGMENT
# # # # #         seg_rules = [r for r in rules
# # # # #                      if r.get("LOB") == lob and r.get("SEGMENT") == json_seg]

# # # # #         candidate_rules = filter_by_insurer(seg_rules, ins)
# # # # #         selected_po = select_po(candidate_rules, p)

# # # # #         if selected_po is None:
# # # # #             p_out = floor_payout(p * 0.90)
# # # # #         else:
# # # # #             p_out = parse_po_to_payout(selected_po, p)

# # # # #     # ── ODISHA OVERRIDE (additional deduction on top) ─────────────────────────
# # # # #     if "ODISHA" in loc:
# # # # #         odisha_rules = [r for r in rules
# # # # #                         if r.get("LOCATION") == "ODISHA"
# # # # #                         and r.get("SEGMENT") == "ALL SEGMENT"
# # # # #                         and str(r.get("INSURER","")).strip().upper() == "ALL COMPANIES"]
# # # # #         ded_po = select_po(odisha_rules, p)
# # # # #         if ded_po:
# # # # #             ded_str = str(ded_po).strip().upper()
# # # # #             m = re.search(r'-(\d+(?:\.\d+)?)%', ded_str)
# # # # #             if m:
# # # # #                 ded = float(m.group(1))
# # # # #                 p_out = floor_payout(p_out - ded)

# # # # #     return max(0.0, p_out)


# # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # #  MAIN
# # # # # # ─────────────────────────────────────────────────────────────────────────────

# # # # # def main():
# # # # #     print("\n" + "="*70)
# # # # #     print("  Payin-Config — Payout Recalculator")
# # # # #     print("="*70)

# # # # #     json_path  = input("\nEnter path to payout_rules.json : ").strip().strip('"')
# # # # #     input_path = input("Enter path to PayinConfig.xlsx  : ").strip().strip('"')
# # # # #     output_path= input("Enter output file path (blank=auto): ").strip().strip('"')

# # # # #     try:
# # # # #         with open(json_path) as f:
# # # # #             rules = json.load(f)
# # # # #         print(f"  Loaded {len(rules)} rules from {json_path}")
# # # # #     except Exception as e:
# # # # #         print(f"\n[ERROR] Failed to load JSON: {e}"); return

# # # # #     if not output_path:
# # # # #         base, ext = os.path.splitext(input_path)
# # # # #         ts = datetime.now().strftime("%Y%m%d_%H%M%S")
# # # # #         output_path = f"{base}_recalculated_{ts}{ext}"

# # # # #     print(f"\n  Reading: {input_path}")
# # # # #     df = pd.read_excel(input_path)
# # # # #     df.columns = [c.strip() for c in df.columns]
# # # # #     total = len(df)
# # # # #     print(f"  Rows   : {total}")

# # # # #     required = ['payin_od_rate','payin_tp_rate','payout_od_rate','payout_tp_rate',
# # # # #                 'sub_product_name','segment','company_code','rto_group_name',
# # # # #                 'vehicle_type_id','from_weightage_kg','to_weightage_kg']
# # # # #     missing = [c for c in required if c not in df.columns]
# # # # #     if missing:
# # # # #         print(f"\n[ERROR] Missing columns: {missing}"); return

# # # # #     new_od, new_tp = [], []
# # # # #     changed_od = changed_tp = processed_od = processed_tp = 0

# # # # #     for _, row in df.iterrows():
# # # # #         sp   = row['sub_product_name']
# # # # #         seg  = row['segment']
# # # # #         ins  = row['company_code']
# # # # #         loc  = row['rto_group_name']
# # # # #         vt   = row['vehicle_type_id']
# # # # #         f_wt = row['from_weightage_kg']
# # # # #         t_wt = row['to_weightage_kg']

# # # # #         # OD
# # # # #         payin_od = row['payin_od_rate']
# # # # #         old_od   = row['payout_od_rate']
# # # # #         if pd.notna(payin_od) and float(payin_od) != 0:
# # # # #             processed_od += 1
# # # # #             calc_od = compute_payout(payin_od, sp, seg, ins, loc, vt, f_wt, t_wt, rules, is_od=True)
# # # # #             new_od.append(calc_od)
# # # # #             if abs(float(old_od) - calc_od) > 0.001:
# # # # #                 changed_od += 1
# # # # #         else:
# # # # #             new_od.append(0.0 if pd.isna(old_od) else old_od)

# # # # #         # TP
# # # # #         payin_tp = row['payin_tp_rate']
# # # # #         old_tp   = row['payout_tp_rate']
# # # # #         if pd.notna(payin_tp) and float(payin_tp) != 0:
# # # # #             processed_tp += 1
# # # # #             calc_tp = compute_payout(payin_tp, sp, seg, ins, loc, vt, f_wt, t_wt, rules, is_od=False)
# # # # #             new_tp.append(calc_tp)
# # # # #             if abs(float(old_tp) - calc_tp) > 0.001:
# # # # #                 changed_tp += 1
# # # # #         else:
# # # # #             new_tp.append(0.0 if pd.isna(old_tp) else old_tp)

# # # # #     df['payout_od_rate'] = new_od
# # # # #     df['payout_tp_rate'] = new_tp
# # # # #     df.to_excel(output_path, index=False)

# # # # #     print(f"\n{'='*70}")
# # # # #     print(f"  COMPLETED")
# # # # #     print(f"  Total rows           : {total}")
# # # # #     print(f"  OD rows recalculated : {processed_od}   changed: {changed_od}")
# # # # #     print(f"  TP rows recalculated : {processed_tp}   changed: {changed_tp}")
# # # # #     print(f"  Output saved to      : {output_path}")
# # # # #     print(f"{'='*70}")

# # # # #     # Sample preview
# # # # #     sample_cols = ['sub_product_name','segment','company_code','rto_group_name',
# # # # #                    'vehicle_type_id','payin_od_rate','payout_od_rate',
# # # # #                    'payin_tp_rate','payout_tp_rate']
# # # # #     sample = df[df['payin_od_rate'] > 0].head(15)[sample_cols]
# # # # #     print("\n  Sample output (first 15 non-zero OD rows):\n")
# # # # #     print(sample.to_string(index=False))
# # # # #     print()


# # # # # if __name__ == "__main__":
# # # # #     main()

# # # # """
# # # # Payin-Config — Payout Recalculator
# # # # ===================================
# # # # Reads PayinConfig.xlsx + master Vehicle Type table (hardcoded from provided data)
# # # # and recomputes payout_od_rate / payout_tp_rate using JSON payout rules.

# # # # Usage:
# # # #     python recalculate_payout.py
# # # # """

# # # # import pandas as pd
# # # # import math
# # # # import os
# # # # import json
# # # # import re
# # # # from datetime import datetime

# # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # #  MASTER — Vehicle Type lookup (from master file Vehicle type worksheet)
# # # # # ─────────────────────────────────────────────────────────────────────────────

# # # # VEHICLE_TYPE_MASTER = {
# # # #     1:  {"vehicle_type": "Agriculture Tractor",          "sub_product_name": "Miscellaneous Vehicle"},
# # # #     2:  {"vehicle_type": "Non Tractor",                  "sub_product_name": "Miscellaneous Vehicle"},
# # # #     3:  {"vehicle_type": "Truck",                        "sub_product_name": "Goods Vehicle"},
# # # #     4:  {"vehicle_type": "Good Carring Tractor",         "sub_product_name": "Goods Vehicle"},
# # # #     5:  {"vehicle_type": "Tanker",                       "sub_product_name": "Goods Vehicle"},
# # # #     6:  {"vehicle_type": "Pickup",                       "sub_product_name": "Goods Vehicle"},
# # # #     7:  {"vehicle_type": "GCV 3W Delivery Van",          "sub_product_name": "Goods Vehicle"},
# # # #     8:  {"vehicle_type": "Taxi_CAB",                     "sub_product_name": "Passenger Vehicle"},
# # # #     9:  {"vehicle_type": "Electric Rikshaw",             "sub_product_name": "Passenger Vehicle"},
# # # #     10: {"vehicle_type": "Tempo Traveller",              "sub_product_name": "Passenger Vehicle"},
# # # #     11: {"vehicle_type": "School Bus",                   "sub_product_name": "Passenger Vehicle"},
# # # #     12: {"vehicle_type": "Passanger Bus",                "sub_product_name": "Passenger Vehicle"},
# # # #     13: {"vehicle_type": "Auto rikshaw",                 "sub_product_name": "Passenger Vehicle"},
# # # #     14: {"vehicle_type": "3W Tipper",                    "sub_product_name": "Goods Vehicle"},
# # # #     15: {"vehicle_type": "PCV 2W",                       "sub_product_name": "Passenger Vehicle"},
# # # #     16: {"vehicle_type": "GCV 2W",                       "sub_product_name": "Goods Vehicle"},
# # # #     17: {"vehicle_type": "TW Scooter",                   "sub_product_name": "Two Wheeler"},
# # # #     18: {"vehicle_type": "TW Bike",                      "sub_product_name": "Two Wheeler"},
# # # #     19: {"vehicle_type": "Private Car",                  "sub_product_name": "Private Car"},
# # # #     20: {"vehicle_type": "TW Electric Bike",             "sub_product_name": "Two Wheeler"},
# # # #     21: {"vehicle_type": "Electric GCV 3W Delivery Van", "sub_product_name": "Goods Vehicle"},
# # # #     22: {"vehicle_type": "Private Car Electric",         "sub_product_name": "Private Car"},
# # # #     23: {"vehicle_type": "Trailer",                      "sub_product_name": "Goods Vehicle"},
# # # #     24: {"vehicle_type": "Electric Pickup",              "sub_product_name": "Goods Vehicle"},
# # # #     25: {"vehicle_type": "Tipper",                       "sub_product_name": "Goods Vehicle"},
# # # #     26: {"vehicle_type": "TW Electric Scooter",          "sub_product_name": "Two Wheeler"},
# # # #     27: {"vehicle_type": "PC Petrol / Electric Hybrid",  "sub_product_name": "Private Car"},
# # # #     28: {"vehicle_type": "Staff Bus",                    "sub_product_name": "Passenger Vehicle"},
# # # #     29: {"vehicle_type": "Agriculture Harvester",        "sub_product_name": "Miscellaneous Vehicle"},
# # # #     30: {"vehicle_type": "Electric PCV 2W",              "sub_product_name": "Passenger Vehicle"},
# # # #     31: {"vehicle_type": "Electric Taxi_CAB",            "sub_product_name": "Passenger Vehicle"},
# # # #     32: {"vehicle_type": "Route Bus",                    "sub_product_name": "Passenger Vehicle"},
# # # # }

# # # # # IDs considered as TAXI
# # # # TAXI_VEHICLE_IDS = {8, 31}  # Taxi_CAB, Electric Taxi_CAB

# # # # # IDs considered as STAFF BUS
# # # # STAFF_BUS_IDS = {28}  # Staff Bus

# # # # # IDs considered as SCHOOL BUS
# # # # SCHOOL_BUS_IDS = {11}  # School Bus

# # # # # IDs considered as BUS (any bus — route/passenger)
# # # # ROUTE_BUS_IDS = {12, 32}  # Passanger Bus, Route Bus

# # # # # IDs considered as GCV 3-Wheeler goods
# # # # GCV_3W_IDS = {7, 14, 21}  # GCV 3W Delivery Van, 3W Tipper, Electric GCV 3W

# # # # # IDs considered as Passenger 3-Wheeler (auto etc.)
# # # # PCV_3W_IDS = {9, 13, 15, 30}  # Electric Rikshaw, Auto rikshaw, PCV 2W, Electric PCV 2W

# # # # # Insurers to match for Upto 2.5 GVW special rule
# # # # SPECIAL_GCV_INSURERS = {"RELIANCE", "SBI"}

# # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # #  FLOOR
# # # # # ─────────────────────────────────────────────────────────────────────────────

# # # # def floor_payout(value):
# # # #     return float(math.floor(float(value)))


# # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # #  PO STRING PARSER
# # # # # ─────────────────────────────────────────────────────────────────────────────

# # # # def parse_po_to_payout(po_str, p):
# # # #     po_str = str(po_str).strip().upper()

# # # #     if re.search(r'\d+%\s*OF PAYIN', po_str):
# # # #         percent = float(re.search(r'(\d+(?:\.\d+)?)%', po_str).group(1))
# # # #         return floor_payout(p * (percent / 100))

# # # #     if "PAYIN + 1" in po_str:
# # # #         return floor_payout(p + 1)

# # # #     if "LESS 2% OF PAYIN" in po_str:
# # # #         return floor_payout(p - 2)

# # # #     # "-3%", "-4%", "-5%"
# # # #     m = re.fullmatch(r'-(\d+(?:\.\d+)?)%', po_str)
# # # #     if m:
# # # #         ded = float(m.group(1))
# # # #         return floor_payout(p - ded)

# # # #     # "21% PO" — fixed payout
# # # #     m = re.search(r'(\d+(?:\.\d+)?)%\s*PO', po_str)
# # # #     if m:
# # # #         return floor_payout(float(m.group(1)))

# # # #     # Fallback
# # # #     return floor_payout(p * 0.90)


# # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # #  RULE SELECTION — slab-based matching on REMARKS
# # # # # ─────────────────────────────────────────────────────────────────────────────

# # # # def select_po(rules_list, p):
# # # #     if not rules_list:
# # # #         return None
# # # #     for r in rules_list:
# # # #         rem = str(r.get("REMARKS", "")).upper().strip()
# # # #         if not rem or rem == "NIL" or rem == "ALL FUEL":
# # # #             return r["PO"]
# # # #         m_below = re.search(r'BELOW\s+(\d+)%', rem)
# # # #         if m_below and p <= float(m_below.group(1)):
# # # #             return r["PO"]
# # # #         m_range = re.search(r'(\d+)%\s+TO\s+(\d+)%', rem)
# # # #         if m_range:
# # # #             lo, hi = float(m_range.group(1)), float(m_range.group(2))
# # # #             if lo <= p <= hi:
# # # #                 return r["PO"]
# # # #         m_above = re.search(r'ABOVE\s+(\d+)%', rem)
# # # #         if m_above and p > float(m_above.group(1)):
# # # #             return r["PO"]
# # # #     return rules_list[0]["PO"]


# # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # #  INSURER MATCHING HELPER
# # # # # ─────────────────────────────────────────────────────────────────────────────

# # # # def insurer_matches(rule_insurer_str, ins):
# # # #     """Check if the row's insurer matches a rule's INSURER field."""
# # # #     ri = str(rule_insurer_str).strip().upper()
# # # #     if ri == "ALL COMPANIES":
# # # #         return True
# # # #     rule_insurers = [x.strip().upper() for x in ri.split(",")]
# # # #     return ins.upper() in rule_insurers


# # # # def filter_by_insurer(rules, ins):
# # # #     """
# # # #     Return best matching rules for given insurer:
# # # #     1. Specific match (not 'All Companies', not 'Rest of Companies')
# # # #     2. 'All Companies'
# # # #     3. 'Rest of Companies'
# # # #     """
# # # #     specific = [r for r in rules
# # # #                 if str(r.get("INSURER","")).strip().upper() not in ("ALL COMPANIES","REST OF COMPANIES")
# # # #                 and insurer_matches(r.get("INSURER",""), ins)]
# # # #     if specific:
# # # #         return specific

# # # #     all_co = [r for r in rules if str(r.get("INSURER","")).strip().upper() == "ALL COMPANIES"]
# # # #     if all_co:
# # # #         return all_co

# # # #     rest = [r for r in rules if str(r.get("INSURER","")).strip().upper() == "REST OF COMPANIES"]
# # # #     return rest


# # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # #  DETERMINE JSON SEGMENT from row data
# # # # # ─────────────────────────────────────────────────────────────────────────────

# # # # def get_json_lob_and_segment(sub_product_name, segment, vehicle_type_id,
# # # #                               from_wt, to_wt, company_code, is_od):
# # # #     """
# # # #     Returns (lob, json_segment) tuple for rule lookup.
# # # #     lob is the JSON LOB string.
# # # #     json_segment is the JSON SEGMENT string.
# # # #     Returns (None, None) if not mappable.
# # # #     """
# # # #     sp = str(sub_product_name).strip()
# # # #     seg = str(segment).strip().upper()
# # # #     vt_id = int(vehicle_type_id) if pd.notna(vehicle_type_id) else 0
# # # #     ins_upper = str(company_code).strip().upper()

# # # #     # ── TWO WHEELER ──────────────────────────────────────────────────────────
# # # #     if sp == "Two Wheeler":
# # # #         lob = "TW"
# # # #         if is_od or "COMP" in seg or "SAOD" in seg or "OD" in seg:
# # # #             return lob, "TW SAOD + COMP"
# # # #         else:  # TP Only
# # # #             return lob, "TW TP"

# # # #     # ── PRIVATE CAR ──────────────────────────────────────────────────────────
# # # #     if sp == "Private Car":
# # # #         lob = "PVT CAR"
# # # #         if is_od or "COMP" in seg or "SAOD" in seg or "OD" in seg:
# # # #             return lob, "PVT CAR COMP + SAOD"
# # # #         else:
# # # #             return lob, "PVT CAR TP"

# # # #     # ── PASSENGER VEHICLE ────────────────────────────────────────────────────
# # # #     if sp == "Passenger Vehicle":
# # # #         # TAXI
# # # #         if vt_id in TAXI_VEHICLE_IDS:
# # # #             return "TAXI", "TAXI"

# # # #         # STAFF BUS
# # # #         if vt_id in STAFF_BUS_IDS:
# # # #             return "BUS", "STAFF BUS"

# # # #         # SCHOOL BUS
# # # #         if vt_id in SCHOOL_BUS_IDS:
# # # #             return "BUS", "SCHOOL BUS"

# # # #         # ROUTE/PASSENGER BUS → treat as STAFF BUS (general BUS)
# # # #         if vt_id in ROUTE_BUS_IDS:
# # # #             return "BUS", "STAFF BUS"

# # # #         # 3-wheelers (Auto, Electric Rikshaw, PCV 2W etc.)
# # # #         if vt_id in PCV_3W_IDS:
# # # #             return "GCV, PCV 3W", "All GVW & PCV 3W, GCV 3W"

# # # #         # Tempo Traveller — treat as staff bus
# # # #         return "BUS", "STAFF BUS"

# # # #     # ── GOODS VEHICLE ────────────────────────────────────────────────────────
# # # #     if sp == "Goods Vehicle":
# # # #         from_w = float(from_wt) if pd.notna(from_wt) else 0
# # # #         to_w   = float(to_wt)   if pd.notna(to_wt)   else 99999

# # # #         # 3-Wheeler goods
# # # #         if vt_id in GCV_3W_IDS:
# # # #             return "GCV, PCV 3W", "All GVW & PCV 3W, GCV 3W"

# # # #         # Upto 2.5T GVW + special insurers
# # # #         if from_w == 0 and to_w <= 2500 and ins_upper in SPECIAL_GCV_INSURERS:
# # # #             return "GCV, PCV 3W", "Upto 2.5 GVW"

# # # #         # Everything else (inc. upto 2.5T with other insurers)
# # # #         return "GCV, PCV 3W", "All GVW & PCV 3W, GCV 3W"

# # # #     # ── MISCELLANEOUS VEHICLE (Tractor, Harvester etc.) ───────────────────────
# # # #     if sp == "Miscellaneous Vehicle":
# # # #         return "MISD", "Misd, Tractor"

# # # #     return None, None


# # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # #  CORE FORMULA ENGINE
# # # # # ─────────────────────────────────────────────────────────────────────────────

# # # # def compute_payout(payin, sub_product_name, segment, company_code,
# # # #                    rto_group_name, vehicle_type_id, from_wt, to_wt,
# # # #                    rules, is_od=True):
# # # #     """
# # # #     Returns (payout_value, explanation_dict).
# # # #     explanation_dict has keys: lob, segment, insurer_matched, po_formula,
# # # #     remarks_slab, odisha_deduction, calculation_note
# # # #     """
# # # #     explanation = {
# # # #         "lob": "",
# # # #         "segment": "",
# # # #         "insurer_matched": "",
# # # #         "po_formula": "",
# # # #         "remarks_slab": "",
# # # #         "odisha_deduction": "",
# # # #         "calculation_note": "",
# # # #     }

# # # #     try:
# # # #         p = float(payin)
# # # #     except (TypeError, ValueError):
# # # #         explanation["calculation_note"] = "Invalid payin value"
# # # #         return 0.0, explanation
# # # #     if p == 0:
# # # #         explanation["calculation_note"] = "Payin is 0 — no payout"
# # # #         return 0.0, explanation

# # # #     ins = str(company_code).strip().upper() if pd.notna(company_code) else ""
# # # #     loc = str(rto_group_name).strip().upper() if pd.notna(rto_group_name) else ""

# # # #     lob, json_seg = get_json_lob_and_segment(
# # # #         sub_product_name, segment, vehicle_type_id, from_wt, to_wt, company_code, is_od
# # # #     )

# # # #     explanation["lob"]     = lob     if lob     else "NOT MAPPED"
# # # #     explanation["segment"] = json_seg if json_seg else "NOT MAPPED"

# # # #     if lob is None:
# # # #         p_out = floor_payout(p * 0.90)
# # # #         explanation["po_formula"]        = "90% of Payin (fallback)"
# # # #         explanation["insurer_matched"]   = "N/A"
# # # #         explanation["remarks_slab"]      = "N/A"
# # # #         explanation["calculation_note"]  = f"No LOB mapping found. Fallback: floor({p} × 0.90) = {p_out}"
# # # #     else:
# # # #         seg_rules      = [r for r in rules if r.get("LOB") == lob and r.get("SEGMENT") == json_seg]
# # # #         candidate_rules = filter_by_insurer(seg_rules, ins)
# # # #         selected_po    = select_po(candidate_rules, p)

# # # #         if candidate_rules:
# # # #             explanation["insurer_matched"] = str(candidate_rules[0].get("INSURER", ""))
# # # #             # Find which slab was picked
# # # #             for r in candidate_rules:
# # # #                 rem = str(r.get("REMARKS","")).upper().strip()
# # # #                 if not rem or rem == "NIL" or rem == "ALL FUEL":
# # # #                     explanation["remarks_slab"] = r.get("REMARKS","NIL")
# # # #                     break
# # # #                 m_below = re.search(r'BELOW\s+(\d+)%', rem)
# # # #                 if m_below and p <= float(m_below.group(1)):
# # # #                     explanation["remarks_slab"] = r.get("REMARKS","")
# # # #                     break
# # # #                 m_range = re.search(r'(\d+)%\s+TO\s+(\d+)%', rem)
# # # #                 if m_range:
# # # #                     lo, hi = float(m_range.group(1)), float(m_range.group(2))
# # # #                     if lo <= p <= hi:
# # # #                         explanation["remarks_slab"] = r.get("REMARKS","")
# # # #                         break
# # # #                 m_above = re.search(r'ABOVE\s+(\d+)%', rem)
# # # #                 if m_above and p > float(m_above.group(1)):
# # # #                     explanation["remarks_slab"] = r.get("REMARKS","")
# # # #                     break
# # # #         else:
# # # #             explanation["insurer_matched"] = "No matching rule"
# # # #             explanation["remarks_slab"]    = "N/A"

# # # #         if selected_po is None:
# # # #             p_out = floor_payout(p * 0.90)
# # # #             explanation["po_formula"]       = "90% of Payin (fallback — no rule matched)"
# # # #             explanation["calculation_note"] = f"Fallback: floor({p} × 0.90) = {p_out}"
# # # #         else:
# # # #             explanation["po_formula"] = selected_po
# # # #             p_out = parse_po_to_payout(selected_po, p)
# # # #             # Build human-readable calculation note
# # # #             po_up = str(selected_po).strip().upper()
# # # #             if re.search(r'\d+%\s*OF PAYIN', po_up):
# # # #                 pct = float(re.search(r'(\d+(?:\.\d+)?)%', po_up).group(1))
# # # #                 explanation["calculation_note"] = f"floor({p} × {pct}/100) = {p_out}"
# # # #             elif "PAYIN + 1" in po_up:
# # # #                 explanation["calculation_note"] = f"floor({p} + 1) = {p_out}"
# # # #             elif "LESS 2% OF PAYIN" in po_up:
# # # #                 explanation["calculation_note"] = f"floor({p} - 2) = {p_out}"
# # # #             elif re.fullmatch(r'-(\d+(?:\.\d+)?)%', po_up):
# # # #                 ded = float(re.search(r'-(\d+(?:\.\d+)?)%', po_up).group(1))
# # # #                 explanation["calculation_note"] = f"floor({p} - {ded}) = {p_out}"
# # # #             elif re.search(r'(\d+(?:\.\d+)?)%\s*PO', po_up):
# # # #                 explanation["calculation_note"] = f"Fixed PO = {p_out}"
# # # #             else:
# # # #                 explanation["calculation_note"] = f"= {p_out}"

# # # #     # ── ODISHA OVERRIDE (additional deduction on top) ─────────────────────────
# # # #     if "ODISHA" in loc:
# # # #         odisha_rules = [r for r in rules
# # # #                         if r.get("LOCATION") == "ODISHA"
# # # #                         and r.get("SEGMENT") == "ALL SEGMENT"
# # # #                         and str(r.get("INSURER","")).strip().upper() == "ALL COMPANIES"]
# # # #         ded_po = select_po(odisha_rules, p)
# # # #         if ded_po:
# # # #             ded_str = str(ded_po).strip().upper()
# # # #             m = re.search(r'-(\d+(?:\.\d+)?)%', ded_str)
# # # #             if m:
# # # #                 ded = float(m.group(1))
# # # #                 p_before = p_out
# # # #                 p_out = floor_payout(p_out - ded)
# # # #                 explanation["odisha_deduction"] = (
# # # #                     f"Odisha override ({ded_po}): floor({p_before} - {ded}) = {p_out}"
# # # #                 )

# # # #     result = max(0.0, p_out)
# # # #     return result, explanation


# # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # #  MAIN
# # # # # ─────────────────────────────────────────────────────────────────────────────

# # # # def main():
# # # #     print("\n" + "="*70)
# # # #     print("  Payin-Config — Payout Recalculator")
# # # #     print("="*70)

# # # #     json_path  = input("\nEnter path to payout_rules.json : ").strip().strip('"')
# # # #     input_path = input("Enter path to PayinConfig.xlsx  : ").strip().strip('"')
# # # #     output_path= input("Enter output file path (blank=auto): ").strip().strip('"')

# # # #     try:
# # # #         with open(json_path) as f:
# # # #             rules = json.load(f)
# # # #         print(f"  Loaded {len(rules)} rules from {json_path}")
# # # #     except Exception as e:
# # # #         print(f"\n[ERROR] Failed to load JSON: {e}"); return

# # # #     if not output_path:
# # # #         base, ext = os.path.splitext(input_path)
# # # #         ts = datetime.now().strftime("%Y%m%d_%H%M%S")
# # # #         output_path = f"{base}_recalculated_{ts}{ext}"

# # # #     print(f"\n  Reading: {input_path}")
# # # #     df = pd.read_excel(input_path)
# # # #     df.columns = [c.strip() for c in df.columns]
# # # #     total = len(df)
# # # #     print(f"  Rows   : {total}")

# # # #     required = ['payin_od_rate','payin_tp_rate','payout_od_rate','payout_tp_rate',
# # # #                 'sub_product_name','segment','company_code','rto_group_name',
# # # #                 'vehicle_type_id','from_weightage_kg','to_weightage_kg']
# # # #     missing = [c for c in required if c not in df.columns]
# # # #     if missing:
# # # #         print(f"\n[ERROR] Missing columns: {missing}"); return

# # # #     new_od, new_tp = [], []
# # # #     changed_od = changed_tp = processed_od = processed_tp = 0

# # # #     # Explanation column lists
# # # #     od_lob, od_seg, od_ins, od_po, od_slab, od_odisha, od_note = [], [], [], [], [], [], []
# # # #     tp_lob, tp_seg, tp_ins, tp_po, tp_slab, tp_odisha, tp_note = [], [], [], [], [], [], []

# # # #     for _, row in df.iterrows():
# # # #         sp   = row['sub_product_name']
# # # #         seg  = row['segment']
# # # #         ins  = row['company_code']
# # # #         loc  = row['rto_group_name']
# # # #         vt   = row['vehicle_type_id']
# # # #         f_wt = row['from_weightage_kg']
# # # #         t_wt = row['to_weightage_kg']

# # # #         # OD
# # # #         payin_od = row['payin_od_rate']
# # # #         old_od   = row['payout_od_rate']
# # # #         if pd.notna(payin_od) and float(payin_od) != 0:
# # # #             processed_od += 1
# # # #             calc_od, expl_od = compute_payout(payin_od, sp, seg, ins, loc, vt, f_wt, t_wt, rules, is_od=True)
# # # #             new_od.append(calc_od)
# # # #             if abs(float(old_od) - calc_od) > 0.001:
# # # #                 changed_od += 1
# # # #         else:
# # # #             calc_od = 0.0 if pd.isna(old_od) else old_od
# # # #             new_od.append(calc_od)
# # # #             expl_od = {"lob":"","segment":"","insurer_matched":"","po_formula":"",
# # # #                        "remarks_slab":"","odisha_deduction":"","calculation_note":"Payin is 0"}
# # # #         od_lob.append(expl_od["lob"])
# # # #         od_seg.append(expl_od["segment"])
# # # #         od_ins.append(expl_od["insurer_matched"])
# # # #         od_po.append(expl_od["po_formula"])
# # # #         od_slab.append(expl_od["remarks_slab"])
# # # #         od_odisha.append(expl_od["odisha_deduction"])
# # # #         od_note.append(expl_od["calculation_note"])

# # # #         # TP
# # # #         payin_tp = row['payin_tp_rate']
# # # #         old_tp   = row['payout_tp_rate']
# # # #         if pd.notna(payin_tp) and float(payin_tp) != 0:
# # # #             processed_tp += 1
# # # #             calc_tp, expl_tp = compute_payout(payin_tp, sp, seg, ins, loc, vt, f_wt, t_wt, rules, is_od=False)
# # # #             new_tp.append(calc_tp)
# # # #             if abs(float(old_tp) - calc_tp) > 0.001:
# # # #                 changed_tp += 1
# # # #         else:
# # # #             calc_tp = 0.0 if pd.isna(old_tp) else old_tp
# # # #             new_tp.append(calc_tp)
# # # #             expl_tp = {"lob":"","segment":"","insurer_matched":"","po_formula":"",
# # # #                        "remarks_slab":"","odisha_deduction":"","calculation_note":"Payin is 0"}
# # # #         tp_lob.append(expl_tp["lob"])
# # # #         tp_seg.append(expl_tp["segment"])
# # # #         tp_ins.append(expl_tp["insurer_matched"])
# # # #         tp_po.append(expl_tp["po_formula"])
# # # #         tp_slab.append(expl_tp["remarks_slab"])
# # # #         tp_odisha.append(expl_tp["odisha_deduction"])
# # # #         tp_note.append(expl_tp["calculation_note"])

# # # #     df['payout_od_rate'] = new_od
# # # #     df['payout_tp_rate'] = new_tp

# # # #     # Append OD explanation columns
# # # #     df['od_rule_lob']            = od_lob
# # # #     df['od_rule_segment']        = od_seg
# # # #     df['od_rule_insurer']        = od_ins
# # # #     df['od_rule_po_formula']     = od_po
# # # #     df['od_rule_slab']           = od_slab
# # # #     df['od_rule_odisha']         = od_odisha
# # # #     df['od_rule_calculation']    = od_note

# # # #     # Append TP explanation columns
# # # #     df['tp_rule_lob']            = tp_lob
# # # #     df['tp_rule_segment']        = tp_seg
# # # #     df['tp_rule_insurer']        = tp_ins
# # # #     df['tp_rule_po_formula']     = tp_po
# # # #     df['tp_rule_slab']           = tp_slab
# # # #     df['tp_rule_odisha']         = tp_odisha
# # # #     df['tp_rule_calculation']    = tp_note

# # # #     df.to_excel(output_path, index=False)

# # # #     print(f"\n{'='*70}")
# # # #     print(f"  COMPLETED")
# # # #     print(f"  Total rows           : {total}")
# # # #     print(f"  OD rows recalculated : {processed_od}   changed: {changed_od}")
# # # #     print(f"  TP rows recalculated : {processed_tp}   changed: {changed_tp}")
# # # #     print(f"  Output saved to      : {output_path}")
# # # #     print(f"{'='*70}")

# # # #     # Sample preview
# # # #     sample_cols = ['sub_product_name','segment','company_code','rto_group_name',
# # # #                    'vehicle_type_id','payin_od_rate','payout_od_rate',
# # # #                    'payin_tp_rate','payout_tp_rate']
# # # #     sample = df[df['payin_od_rate'] > 0].head(15)[sample_cols]
# # # #     print("\n  Sample output (first 15 non-zero OD rows):\n")
# # # #     print(sample.to_string(index=False))
# # # #     print()


# # # # if __name__ == "__main__":
# # # #     main()


# # # """
# # # Payin-Config — Payout Recalculator
# # # ===================================
# # # Reads PayinConfig.xlsx + master Vehicle Type table (hardcoded from provided data)
# # # and recomputes payout_od_rate / payout_tp_rate using JSON payout rules.

# # # Usage:
# # #     python recalculate_payout.py
# # # """

# # # import pandas as pd
# # # import math
# # # import os
# # # import json
# # # import re
# # # from datetime import datetime

# # # # ─────────────────────────────────────────────────────────────────────────────
# # # #  MASTER — Vehicle Type lookup (from master file Vehicle type worksheet)
# # # # ─────────────────────────────────────────────────────────────────────────────

# # # VEHICLE_TYPE_MASTER = {
# # #     1:  {"vehicle_type": "Agriculture Tractor",          "sub_product_name": "Miscellaneous Vehicle"},
# # #     2:  {"vehicle_type": "Non Tractor",                  "sub_product_name": "Miscellaneous Vehicle"},
# # #     3:  {"vehicle_type": "Truck",                        "sub_product_name": "Goods Vehicle"},
# # #     4:  {"vehicle_type": "Good Carring Tractor",         "sub_product_name": "Goods Vehicle"},
# # #     5:  {"vehicle_type": "Tanker",                       "sub_product_name": "Goods Vehicle"},
# # #     6:  {"vehicle_type": "Pickup",                       "sub_product_name": "Goods Vehicle"},
# # #     7:  {"vehicle_type": "GCV 3W Delivery Van",          "sub_product_name": "Goods Vehicle"},
# # #     8:  {"vehicle_type": "Taxi_CAB",                     "sub_product_name": "Passenger Vehicle"},
# # #     9:  {"vehicle_type": "Electric Rikshaw",             "sub_product_name": "Passenger Vehicle"},
# # #     10: {"vehicle_type": "Tempo Traveller",              "sub_product_name": "Passenger Vehicle"},
# # #     11: {"vehicle_type": "School Bus",                   "sub_product_name": "Passenger Vehicle"},
# # #     12: {"vehicle_type": "Passanger Bus",                "sub_product_name": "Passenger Vehicle"},
# # #     13: {"vehicle_type": "Auto rikshaw",                 "sub_product_name": "Passenger Vehicle"},
# # #     14: {"vehicle_type": "3W Tipper",                    "sub_product_name": "Goods Vehicle"},
# # #     15: {"vehicle_type": "PCV 2W",                       "sub_product_name": "Passenger Vehicle"},
# # #     16: {"vehicle_type": "GCV 2W",                       "sub_product_name": "Goods Vehicle"},
# # #     17: {"vehicle_type": "TW Scooter",                   "sub_product_name": "Two Wheeler"},
# # #     18: {"vehicle_type": "TW Bike",                      "sub_product_name": "Two Wheeler"},
# # #     19: {"vehicle_type": "Private Car",                  "sub_product_name": "Private Car"},
# # #     20: {"vehicle_type": "TW Electric Bike",             "sub_product_name": "Two Wheeler"},
# # #     21: {"vehicle_type": "Electric GCV 3W Delivery Van", "sub_product_name": "Goods Vehicle"},
# # #     22: {"vehicle_type": "Private Car Electric",         "sub_product_name": "Private Car"},
# # #     23: {"vehicle_type": "Trailer",                      "sub_product_name": "Goods Vehicle"},
# # #     24: {"vehicle_type": "Electric Pickup",              "sub_product_name": "Goods Vehicle"},
# # #     25: {"vehicle_type": "Tipper",                       "sub_product_name": "Goods Vehicle"},
# # #     26: {"vehicle_type": "TW Electric Scooter",          "sub_product_name": "Two Wheeler"},
# # #     27: {"vehicle_type": "PC Petrol / Electric Hybrid",  "sub_product_name": "Private Car"},
# # #     28: {"vehicle_type": "Staff Bus",                    "sub_product_name": "Passenger Vehicle"},
# # #     29: {"vehicle_type": "Agriculture Harvester",        "sub_product_name": "Miscellaneous Vehicle"},
# # #     30: {"vehicle_type": "Electric PCV 2W",              "sub_product_name": "Passenger Vehicle"},
# # #     31: {"vehicle_type": "Electric Taxi_CAB",            "sub_product_name": "Passenger Vehicle"},
# # #     32: {"vehicle_type": "Route Bus",                    "sub_product_name": "Passenger Vehicle"},
# # # }

# # # # IDs considered as TAXI
# # # TAXI_VEHICLE_IDS = {8, 31}  # Taxi_CAB, Electric Taxi_CAB

# # # # IDs considered as STAFF BUS
# # # STAFF_BUS_IDS = {28}  # Staff Bus

# # # # IDs considered as SCHOOL BUS
# # # SCHOOL_BUS_IDS = {11}  # School Bus

# # # # IDs considered as BUS (any bus — route/passenger)
# # # ROUTE_BUS_IDS = {12, 32}  # Passanger Bus, Route Bus

# # # # IDs considered as GCV 3-Wheeler goods
# # # GCV_3W_IDS = {7, 14, 21}  # GCV 3W Delivery Van, 3W Tipper, Electric GCV 3W

# # # # IDs considered as Passenger 3-Wheeler (auto etc.)
# # # PCV_3W_IDS = {9, 13, 15, 30}  # Electric Rikshaw, Auto rikshaw, PCV 2W, Electric PCV 2W

# # # # Insurers to match for Upto 2.5 GVW special rule
# # # SPECIAL_GCV_INSURERS = {"RELIANCE", "SBI"}

# # # # ─────────────────────────────────────────────────────────────────────────────
# # # #  FLOOR
# # # # ─────────────────────────────────────────────────────────────────────────────

# # # def floor_payout(value):
# # #     return float(math.floor(float(value)))


# # # # ─────────────────────────────────────────────────────────────────────────────
# # # #  PO STRING PARSER
# # # # ─────────────────────────────────────────────────────────────────────────────

# # # def parse_po_to_payout(po_str, p):
# # #     po_str = str(po_str).strip().upper()

# # #     if re.search(r'\d+%\s*OF PAYIN', po_str):
# # #         percent = float(re.search(r'(\d+(?:\.\d+)?)%', po_str).group(1))
# # #         return floor_payout(p * (percent / 100))

# # #     if "PAYIN + 1" in po_str:
# # #         return floor_payout(p + 1)

# # #     if "LESS 2% OF PAYIN" in po_str:
# # #         return floor_payout(p - 2)

# # #     # "-3%", "-4%", "-5%"
# # #     m = re.fullmatch(r'-(\d+(?:\.\d+)?)%', po_str)
# # #     if m:
# # #         ded = float(m.group(1))
# # #         return floor_payout(p - ded)

# # #     # "21% PO" — fixed payout
# # #     m = re.search(r'(\d+(?:\.\d+)?)%\s*PO', po_str)
# # #     if m:
# # #         return floor_payout(float(m.group(1)))

# # #     # Fallback
# # #     return floor_payout(p * 0.90)


# # # # ─────────────────────────────────────────────────────────────────────────────
# # # #  RULE SELECTION — slab-based matching on REMARKS
# # # # ─────────────────────────────────────────────────────────────────────────────

# # # def select_po(rules_list, p):
# # #     if not rules_list:
# # #         return None
# # #     for r in rules_list:
# # #         rem = str(r.get("REMARKS", "")).upper().strip()
# # #         if not rem or rem == "NIL" or rem == "ALL FUEL":
# # #             return r["PO"]
# # #         m_below = re.search(r'BELOW\s+(\d+)%', rem)
# # #         if m_below and p <= float(m_below.group(1)):
# # #             return r["PO"]
# # #         m_range = re.search(r'(\d+)%\s+TO\s+(\d+)%', rem)
# # #         if m_range:
# # #             lo, hi = float(m_range.group(1)), float(m_range.group(2))
# # #             if lo <= p <= hi:
# # #                 return r["PO"]
# # #         m_above = re.search(r'ABOVE\s+(\d+)%', rem)
# # #         if m_above and p > float(m_above.group(1)):
# # #             return r["PO"]
# # #     return rules_list[0]["PO"]


# # # # ─────────────────────────────────────────────────────────────────────────────
# # # #  INSURER MATCHING HELPER
# # # # ─────────────────────────────────────────────────────────────────────────────

# # # def insurer_matches(rule_insurer_str, ins):
# # #     """
# # #     Check if the row's insurer matches a rule's INSURER field.

# # #     Handles entries like "Tata- Comp" where the rule token contains the
# # #     insurer name plus extra words/punctuation (e.g. product suffix).

# # #     Strategy per comma-separated token in the rule:
# # #       1. Exact match after normalisation.
# # #       2. First word of token equals insurer  →  "TATA- COMP" first word = "TATA".
# # #       3. Insurer is a substring of the token as a final fallback.
# # #     """
# # #     ri = str(rule_insurer_str).strip().upper()
# # #     if ri == "ALL COMPANIES":
# # #         return True

# # #     ins_norm = ins.strip().upper()
# # #     for token in [x.strip().upper() for x in ri.split(",")]:
# # #         if ins_norm == token:
# # #             return True
# # #         # Strip punctuation so "TATA- COMP" → ["TATA", "COMP"]
# # #         token_words = re.sub(r'[^A-Z0-9 ]', ' ', token).split()
# # #         if token_words and token_words[0] == ins_norm:
# # #             return True
# # #         if ins_norm in token:
# # #             return True

# # #     return False


# # # def filter_by_insurer(rules, ins):
# # #     """
# # #     Return best matching rules for given insurer:
# # #     1. Specific match (not 'All Companies', not 'Rest of Companies')
# # #     2. 'All Companies'
# # #     3. 'Rest of Companies'
# # #     """
# # #     specific = [r for r in rules
# # #                 if str(r.get("INSURER","")).strip().upper() not in ("ALL COMPANIES","REST OF COMPANIES")
# # #                 and insurer_matches(r.get("INSURER",""), ins)]
# # #     if specific:
# # #         return specific

# # #     all_co = [r for r in rules if str(r.get("INSURER","")).strip().upper() == "ALL COMPANIES"]
# # #     if all_co:
# # #         return all_co

# # #     rest = [r for r in rules if str(r.get("INSURER","")).strip().upper() == "REST OF COMPANIES"]
# # #     return rest


# # # # ─────────────────────────────────────────────────────────────────────────────
# # # #  DETERMINE JSON SEGMENT from row data
# # # # ─────────────────────────────────────────────────────────────────────────────

# # # def get_json_lob_and_segment(sub_product_name, segment, vehicle_type_id,
# # #                               from_wt, to_wt, company_code, is_od):
# # #     """
# # #     Returns (lob, json_segment) tuple for rule lookup.
# # #     lob is the JSON LOB string.
# # #     json_segment is the JSON SEGMENT string.
# # #     Returns (None, None) if not mappable.
# # #     """
# # #     sp = str(sub_product_name).strip()
# # #     seg = str(segment).strip().upper()
# # #     vt_id = int(vehicle_type_id) if pd.notna(vehicle_type_id) else 0
# # #     ins_upper = str(company_code).strip().upper()

# # #     # ── TWO WHEELER ──────────────────────────────────────────────────────────
# # #     if sp == "Two Wheeler":
# # #         lob = "TW"
# # #         if is_od or "COMP" in seg or "SAOD" in seg or "OD" in seg:
# # #             return lob, "TW SAOD + COMP"
# # #         else:  # TP Only
# # #             return lob, "TW TP"

# # #     # ── PRIVATE CAR ──────────────────────────────────────────────────────────
# # #     if sp == "Private Car":
# # #         lob = "PVT CAR"
# # #         if is_od or "COMP" in seg or "SAOD" in seg or "OD" in seg:
# # #             return lob, "PVT CAR COMP + SAOD"
# # #         else:
# # #             return lob, "PVT CAR TP"

# # #     # ── PASSENGER VEHICLE ────────────────────────────────────────────────────
# # #     if sp == "Passenger Vehicle":
# # #         # TAXI
# # #         if vt_id in TAXI_VEHICLE_IDS:
# # #             return "TAXI", "TAXI"

# # #         # STAFF BUS
# # #         if vt_id in STAFF_BUS_IDS:
# # #             return "BUS", "STAFF BUS"

# # #         # SCHOOL BUS
# # #         if vt_id in SCHOOL_BUS_IDS:
# # #             return "BUS", "SCHOOL BUS"

# # #         # ROUTE/PASSENGER BUS → treat as STAFF BUS (general BUS)
# # #         if vt_id in ROUTE_BUS_IDS:
# # #             return "BUS", "STAFF BUS"

# # #         # 3-wheelers (Auto, Electric Rikshaw, PCV 2W etc.)
# # #         if vt_id in PCV_3W_IDS:
# # #             return "GCV, PCV 3W", "All GVW & PCV 3W, GCV 3W"

# # #         # Tempo Traveller — treat as staff bus
# # #         return "BUS", "STAFF BUS"

# # #     # ── GOODS VEHICLE ────────────────────────────────────────────────────────
# # #     if sp == "Goods Vehicle":
# # #         from_w = float(from_wt) if pd.notna(from_wt) else 0
# # #         to_w   = float(to_wt)   if pd.notna(to_wt)   else 99999

# # #         # 3-Wheeler goods
# # #         if vt_id in GCV_3W_IDS:
# # #             return "GCV, PCV 3W", "All GVW & PCV 3W, GCV 3W"

# # #         # Upto 2.5T GVW + special insurers
# # #         if from_w == 0 and to_w <= 2500 and ins_upper in SPECIAL_GCV_INSURERS:
# # #             return "GCV, PCV 3W", "Upto 2.5 GVW"

# # #         # Everything else (inc. upto 2.5T with other insurers)
# # #         return "GCV, PCV 3W", "All GVW & PCV 3W, GCV 3W"

# # #     # ── MISCELLANEOUS VEHICLE (Tractor, Harvester etc.) ───────────────────────
# # #     if sp == "Miscellaneous Vehicle":
# # #         return "MISD", "Misd, Tractor"

# # #     return None, None


# # # # ─────────────────────────────────────────────────────────────────────────────
# # # #  CORE FORMULA ENGINE
# # # # ─────────────────────────────────────────────────────────────────────────────

# # # def compute_payout(payin, sub_product_name, segment, company_code,
# # #                    rto_group_name, vehicle_type_id, from_wt, to_wt,
# # #                    rules, is_od=True):
# # #     """
# # #     Returns (payout_value, explanation_dict).
# # #     explanation_dict has keys: lob, segment, insurer_matched, po_formula,
# # #     remarks_slab, odisha_deduction, calculation_note
# # #     """
# # #     explanation = {
# # #         "lob": "",
# # #         "segment": "",
# # #         "insurer_matched": "",
# # #         "po_formula": "",
# # #         "remarks_slab": "",
# # #         "odisha_deduction": "",
# # #         "calculation_note": "",
# # #     }

# # #     try:
# # #         p = float(payin)
# # #     except (TypeError, ValueError):
# # #         explanation["calculation_note"] = "Invalid payin value"
# # #         return 0.0, explanation
# # #     if p == 0:
# # #         explanation["calculation_note"] = "Payin is 0 — no payout"
# # #         return 0.0, explanation

# # #     ins = str(company_code).strip().upper() if pd.notna(company_code) else ""
# # #     loc = str(rto_group_name).strip().upper() if pd.notna(rto_group_name) else ""

# # #     lob, json_seg = get_json_lob_and_segment(
# # #         sub_product_name, segment, vehicle_type_id, from_wt, to_wt, company_code, is_od
# # #     )

# # #     explanation["lob"]     = lob     if lob     else "NOT MAPPED"
# # #     explanation["segment"] = json_seg if json_seg else "NOT MAPPED"

# # #     if lob is None:
# # #         p_out = floor_payout(p * 0.90)
# # #         explanation["po_formula"]        = "90% of Payin (fallback)"
# # #         explanation["insurer_matched"]   = "N/A"
# # #         explanation["remarks_slab"]      = "N/A"
# # #         explanation["calculation_note"]  = f"No LOB mapping found. Fallback: floor({p} × 0.90) = {p_out}"
# # #     else:
# # #         seg_rules      = [r for r in rules if r.get("LOB") == lob and r.get("SEGMENT") == json_seg]
# # #         candidate_rules = filter_by_insurer(seg_rules, ins)
# # #         selected_po    = select_po(candidate_rules, p)

# # #         if candidate_rules:
# # #             explanation["insurer_matched"] = str(candidate_rules[0].get("INSURER", ""))
# # #             # Find which slab was picked
# # #             for r in candidate_rules:
# # #                 rem = str(r.get("REMARKS","")).upper().strip()
# # #                 if not rem or rem == "NIL" or rem == "ALL FUEL":
# # #                     explanation["remarks_slab"] = r.get("REMARKS","NIL")
# # #                     break
# # #                 m_below = re.search(r'BELOW\s+(\d+)%', rem)
# # #                 if m_below and p <= float(m_below.group(1)):
# # #                     explanation["remarks_slab"] = r.get("REMARKS","")
# # #                     break
# # #                 m_range = re.search(r'(\d+)%\s+TO\s+(\d+)%', rem)
# # #                 if m_range:
# # #                     lo, hi = float(m_range.group(1)), float(m_range.group(2))
# # #                     if lo <= p <= hi:
# # #                         explanation["remarks_slab"] = r.get("REMARKS","")
# # #                         break
# # #                 m_above = re.search(r'ABOVE\s+(\d+)%', rem)
# # #                 if m_above and p > float(m_above.group(1)):
# # #                     explanation["remarks_slab"] = r.get("REMARKS","")
# # #                     break
# # #         else:
# # #             explanation["insurer_matched"] = "No matching rule"
# # #             explanation["remarks_slab"]    = "N/A"

# # #         if selected_po is None:
# # #             p_out = floor_payout(p * 0.90)
# # #             explanation["po_formula"]       = "90% of Payin (fallback — no rule matched)"
# # #             explanation["calculation_note"] = f"Fallback: floor({p} × 0.90) = {p_out}"
# # #         else:
# # #             explanation["po_formula"] = selected_po
# # #             p_out = parse_po_to_payout(selected_po, p)
# # #             # Build human-readable calculation note
# # #             po_up = str(selected_po).strip().upper()
# # #             if re.search(r'\d+%\s*OF PAYIN', po_up):
# # #                 pct = float(re.search(r'(\d+(?:\.\d+)?)%', po_up).group(1))
# # #                 explanation["calculation_note"] = f"floor({p} × {pct}/100) = {p_out}"
# # #             elif "PAYIN + 1" in po_up:
# # #                 explanation["calculation_note"] = f"floor({p} + 1) = {p_out}"
# # #             elif "LESS 2% OF PAYIN" in po_up:
# # #                 explanation["calculation_note"] = f"floor({p} - 2) = {p_out}"
# # #             elif re.fullmatch(r'-(\d+(?:\.\d+)?)%', po_up):
# # #                 ded = float(re.search(r'-(\d+(?:\.\d+)?)%', po_up).group(1))
# # #                 explanation["calculation_note"] = f"floor({p} - {ded}) = {p_out}"
# # #             elif re.search(r'(\d+(?:\.\d+)?)%\s*PO', po_up):
# # #                 explanation["calculation_note"] = f"Fixed PO = {p_out}"
# # #             else:
# # #                 explanation["calculation_note"] = f"= {p_out}"

# # #     # ── ODISHA OVERRIDE (additional deduction on top) ─────────────────────────
# # #     if "ODISHA" in loc:
# # #         odisha_rules = [r for r in rules
# # #                         if r.get("LOCATION") == "ODISHA"
# # #                         and r.get("SEGMENT") == "ALL SEGMENT"
# # #                         and str(r.get("INSURER","")).strip().upper() == "ALL COMPANIES"]
# # #         ded_po = select_po(odisha_rules, p)
# # #         if ded_po:
# # #             ded_str = str(ded_po).strip().upper()
# # #             m = re.search(r'-(\d+(?:\.\d+)?)%', ded_str)
# # #             if m:
# # #                 ded = float(m.group(1))
# # #                 p_before = p_out
# # #                 p_out = floor_payout(p_out - ded)
# # #                 explanation["odisha_deduction"] = (
# # #                     f"Odisha override ({ded_po}): floor({p_before} - {ded}) = {p_out}"
# # #                 )

# # #     result = max(0.0, p_out)
# # #     return result, explanation


# # # # ─────────────────────────────────────────────────────────────────────────────
# # # #  MAIN
# # # # ─────────────────────────────────────────────────────────────────────────────

# # # def main():
# # #     print("\n" + "="*70)
# # #     print("  Payin-Config — Payout Recalculator")
# # #     print("="*70)

# # #     json_path  = input("\nEnter path to payout_rules.json : ").strip().strip('"')
# # #     input_path = input("Enter path to PayinConfig.xlsx  : ").strip().strip('"')
# # #     output_path= input("Enter output file path (blank=auto): ").strip().strip('"')

# # #     try:
# # #         with open(json_path) as f:
# # #             rules = json.load(f)
# # #         print(f"  Loaded {len(rules)} rules from {json_path}")
# # #     except Exception as e:
# # #         print(f"\n[ERROR] Failed to load JSON: {e}"); return

# # #     if not output_path:
# # #         base, ext = os.path.splitext(input_path)
# # #         ts = datetime.now().strftime("%Y%m%d_%H%M%S")
# # #         output_path = f"{base}_recalculated_{ts}{ext}"

# # #     print(f"\n  Reading: {input_path}")
# # #     df = pd.read_excel(input_path)
# # #     df.columns = [c.strip() for c in df.columns]
# # #     total = len(df)
# # #     print(f"  Rows   : {total}")

# # #     required = ['payin_od_rate','payin_tp_rate','payout_od_rate','payout_tp_rate',
# # #                 'sub_product_name','segment','company_code','rto_group_name',
# # #                 'vehicle_type_id','from_weightage_kg','to_weightage_kg']
# # #     missing = [c for c in required if c not in df.columns]
# # #     if missing:
# # #         print(f"\n[ERROR] Missing columns: {missing}"); return

# # #     new_od, new_tp = [], []
# # #     changed_od = changed_tp = processed_od = processed_tp = 0

# # #     # Explanation column lists
# # #     od_lob, od_seg, od_ins, od_po, od_slab, od_odisha, od_note = [], [], [], [], [], [], []
# # #     tp_lob, tp_seg, tp_ins, tp_po, tp_slab, tp_odisha, tp_note = [], [], [], [], [], [], []

# # #     for _, row in df.iterrows():
# # #         sp   = row['sub_product_name']
# # #         seg  = row['segment']
# # #         ins  = row['company_code']
# # #         loc  = row['rto_group_name']
# # #         vt   = row['vehicle_type_id']
# # #         f_wt = row['from_weightage_kg']
# # #         t_wt = row['to_weightage_kg']

# # #         # OD
# # #         payin_od = row['payin_od_rate']
# # #         old_od   = row['payout_od_rate']
# # #         if pd.notna(payin_od) and float(payin_od) != 0:
# # #             processed_od += 1
# # #             calc_od, expl_od = compute_payout(payin_od, sp, seg, ins, loc, vt, f_wt, t_wt, rules, is_od=True)
# # #             new_od.append(calc_od)
# # #             if abs(float(old_od) - calc_od) > 0.001:
# # #                 changed_od += 1
# # #         else:
# # #             calc_od = 0.0 if pd.isna(old_od) else old_od
# # #             new_od.append(calc_od)
# # #             expl_od = {"lob":"","segment":"","insurer_matched":"","po_formula":"",
# # #                        "remarks_slab":"","odisha_deduction":"","calculation_note":"Payin is 0"}
# # #         od_lob.append(expl_od["lob"])
# # #         od_seg.append(expl_od["segment"])
# # #         od_ins.append(expl_od["insurer_matched"])
# # #         od_po.append(expl_od["po_formula"])
# # #         od_slab.append(expl_od["remarks_slab"])
# # #         od_odisha.append(expl_od["odisha_deduction"])
# # #         od_note.append(expl_od["calculation_note"])

# # #         # TP
# # #         payin_tp = row['payin_tp_rate']
# # #         old_tp   = row['payout_tp_rate']
# # #         if pd.notna(payin_tp) and float(payin_tp) != 0:
# # #             processed_tp += 1
# # #             calc_tp, expl_tp = compute_payout(payin_tp, sp, seg, ins, loc, vt, f_wt, t_wt, rules, is_od=False)
# # #             new_tp.append(calc_tp)
# # #             if abs(float(old_tp) - calc_tp) > 0.001:
# # #                 changed_tp += 1
# # #         else:
# # #             calc_tp = 0.0 if pd.isna(old_tp) else old_tp
# # #             new_tp.append(calc_tp)
# # #             expl_tp = {"lob":"","segment":"","insurer_matched":"","po_formula":"",
# # #                        "remarks_slab":"","odisha_deduction":"","calculation_note":"Payin is 0"}
# # #         tp_lob.append(expl_tp["lob"])
# # #         tp_seg.append(expl_tp["segment"])
# # #         tp_ins.append(expl_tp["insurer_matched"])
# # #         tp_po.append(expl_tp["po_formula"])
# # #         tp_slab.append(expl_tp["remarks_slab"])
# # #         tp_odisha.append(expl_tp["odisha_deduction"])
# # #         tp_note.append(expl_tp["calculation_note"])

# # #     df['payout_od_rate'] = new_od
# # #     df['payout_tp_rate'] = new_tp

# # #     # Append OD explanation columns
# # #     df['od_rule_lob']            = od_lob
# # #     df['od_rule_segment']        = od_seg
# # #     df['od_rule_insurer']        = od_ins
# # #     df['od_rule_po_formula']     = od_po
# # #     df['od_rule_slab']           = od_slab
# # #     df['od_rule_odisha']         = od_odisha
# # #     df['od_rule_calculation']    = od_note

# # #     # Append TP explanation columns
# # #     df['tp_rule_lob']            = tp_lob
# # #     df['tp_rule_segment']        = tp_seg
# # #     df['tp_rule_insurer']        = tp_ins
# # #     df['tp_rule_po_formula']     = tp_po
# # #     df['tp_rule_slab']           = tp_slab
# # #     df['tp_rule_odisha']         = tp_odisha
# # #     df['tp_rule_calculation']    = tp_note

# # #     df.to_excel(output_path, index=False)

# # #     print(f"\n{'='*70}")
# # #     print(f"  COMPLETED")
# # #     print(f"  Total rows           : {total}")
# # #     print(f"  OD rows recalculated : {processed_od}   changed: {changed_od}")
# # #     print(f"  TP rows recalculated : {processed_tp}   changed: {changed_tp}")
# # #     print(f"  Output saved to      : {output_path}")
# # #     print(f"{'='*70}")

# # #     # Sample preview
# # #     sample_cols = ['sub_product_name','segment','company_code','rto_group_name',
# # #                    'vehicle_type_id','payin_od_rate','payout_od_rate',
# # #                    'payin_tp_rate','payout_tp_rate']
# # #     sample = df[df['payin_od_rate'] > 0].head(15)[sample_cols]
# # #     print("\n  Sample output (first 15 non-zero OD rows):\n")
# # #     print(sample.to_string(index=False))
# # #     print()


# # # if __name__ == "__main__":
# # #     main()


# # # """
# # # Payin-Config — Payout Recalculator
# # # ===================================
# # # Reads PayinConfig.xlsx + master Vehicle Type table (hardcoded from provided data)
# # # and recomputes payout_od_rate / payout_tp_rate using JSON payout rules.

# # # Usage:
# # #     python recalculate_payout.py
# # # """

# # # import pandas as pd
# # # import math
# # # import os
# # # import json
# # # import re
# # # from datetime import datetime

# # # # ─────────────────────────────────────────────────────────────────────────────
# # # #  MASTER — Vehicle Type lookup (from master file Vehicle type worksheet)
# # # # ─────────────────────────────────────────────────────────────────────────────

# # # VEHICLE_TYPE_MASTER = {
# # #     1:  {"vehicle_type": "Agriculture Tractor",          "sub_product_name": "Miscellaneous Vehicle"},
# # #     2:  {"vehicle_type": "Non Tractor",                  "sub_product_name": "Miscellaneous Vehicle"},
# # #     3:  {"vehicle_type": "Truck",                        "sub_product_name": "Goods Vehicle"},
# # #     4:  {"vehicle_type": "Good Carring Tractor",         "sub_product_name": "Goods Vehicle"},
# # #     5:  {"vehicle_type": "Tanker",                       "sub_product_name": "Goods Vehicle"},
# # #     6:  {"vehicle_type": "Pickup",                       "sub_product_name": "Goods Vehicle"},
# # #     7:  {"vehicle_type": "GCV 3W Delivery Van",          "sub_product_name": "Goods Vehicle"},
# # #     8:  {"vehicle_type": "Taxi_CAB",                     "sub_product_name": "Passenger Vehicle"},
# # #     9:  {"vehicle_type": "Electric Rikshaw",             "sub_product_name": "Passenger Vehicle"},
# # #     10: {"vehicle_type": "Tempo Traveller",              "sub_product_name": "Passenger Vehicle"},
# # #     11: {"vehicle_type": "School Bus",                   "sub_product_name": "Passenger Vehicle"},
# # #     12: {"vehicle_type": "Passanger Bus",                "sub_product_name": "Passenger Vehicle"},
# # #     13: {"vehicle_type": "Auto rikshaw",                 "sub_product_name": "Passenger Vehicle"},
# # #     14: {"vehicle_type": "3W Tipper",                    "sub_product_name": "Goods Vehicle"},
# # #     15: {"vehicle_type": "PCV 2W",                       "sub_product_name": "Passenger Vehicle"},
# # #     16: {"vehicle_type": "GCV 2W",                       "sub_product_name": "Goods Vehicle"},
# # #     17: {"vehicle_type": "TW Scooter",                   "sub_product_name": "Two Wheeler"},
# # #     18: {"vehicle_type": "TW Bike",                      "sub_product_name": "Two Wheeler"},
# # #     19: {"vehicle_type": "Private Car",                  "sub_product_name": "Private Car"},
# # #     20: {"vehicle_type": "TW Electric Bike",             "sub_product_name": "Two Wheeler"},
# # #     21: {"vehicle_type": "Electric GCV 3W Delivery Van", "sub_product_name": "Goods Vehicle"},
# # #     22: {"vehicle_type": "Private Car Electric",         "sub_product_name": "Private Car"},
# # #     23: {"vehicle_type": "Trailer",                      "sub_product_name": "Goods Vehicle"},
# # #     24: {"vehicle_type": "Electric Pickup",              "sub_product_name": "Goods Vehicle"},
# # #     25: {"vehicle_type": "Tipper",                       "sub_product_name": "Goods Vehicle"},
# # #     26: {"vehicle_type": "TW Electric Scooter",          "sub_product_name": "Two Wheeler"},
# # #     27: {"vehicle_type": "PC Petrol / Electric Hybrid",  "sub_product_name": "Private Car"},
# # #     28: {"vehicle_type": "Staff Bus",                    "sub_product_name": "Passenger Vehicle"},
# # #     29: {"vehicle_type": "Agriculture Harvester",        "sub_product_name": "Miscellaneous Vehicle"},
# # #     30: {"vehicle_type": "Electric PCV 2W",              "sub_product_name": "Passenger Vehicle"},
# # #     31: {"vehicle_type": "Electric Taxi_CAB",            "sub_product_name": "Passenger Vehicle"},
# # #     32: {"vehicle_type": "Route Bus",                    "sub_product_name": "Passenger Vehicle"},
# # # }

# # # # IDs considered as TAXI
# # # TAXI_VEHICLE_IDS = {8, 31}  # Taxi_CAB, Electric Taxi_CAB

# # # # IDs considered as STAFF BUS
# # # STAFF_BUS_IDS = {28}  # Staff Bus

# # # # IDs considered as SCHOOL BUS
# # # SCHOOL_BUS_IDS = {11}  # School Bus

# # # # IDs considered as BUS (any bus — route/passenger)
# # # ROUTE_BUS_IDS = {12, 32}  # Passanger Bus, Route Bus

# # # # IDs considered as GCV 3-Wheeler goods
# # # GCV_3W_IDS = {7, 14, 21}  # GCV 3W Delivery Van, 3W Tipper, Electric GCV 3W

# # # # IDs considered as Passenger 3-Wheeler (auto etc.)
# # # PCV_3W_IDS = {9, 13, 15, 30}  # Electric Rikshaw, Auto rikshaw, PCV 2W, Electric PCV 2W

# # # # Insurers to match for Upto 2.5 GVW special rule
# # # SPECIAL_GCV_INSURERS = {"RELIANCE", "SBI"}

# # # # ─────────────────────────────────────────────────────────────────────────────
# # # #  FLOOR
# # # # ─────────────────────────────────────────────────────────────────────────────

# # # def floor_payout(value):
# # #     return float(math.floor(float(value)))


# # # # ─────────────────────────────────────────────────────────────────────────────
# # # #  PO STRING PARSER
# # # # ─────────────────────────────────────────────────────────────────────────────

# # # def parse_po_to_payout(po_str, p):
# # #     po_str = str(po_str).strip().upper()

# # #     if re.search(r'\d+%\s*OF PAYIN', po_str):
# # #         percent = float(re.search(r'(\d+(?:\.\d+)?)%', po_str).group(1))
# # #         return floor_payout(p * (percent / 100))

# # #     if "PAYIN + 1" in po_str:
# # #         return floor_payout(p + 1)

# # #     if "LESS 2% OF PAYIN" in po_str:
# # #         return floor_payout(p - 2)

# # #     # "-3%", "-4%", "-5%"
# # #     m = re.fullmatch(r'-(\d+(?:\.\d+)?)%', po_str)
# # #     if m:
# # #         ded = float(m.group(1))
# # #         return floor_payout(p - ded)

# # #     # "21% PO" — fixed payout
# # #     m = re.search(r'(\d+(?:\.\d+)?)%\s*PO', po_str)
# # #     if m:
# # #         return floor_payout(float(m.group(1)))

# # #     # Fallback
# # #     return floor_payout(p * 0.90)


# # # # ─────────────────────────────────────────────────────────────────────────────
# # # #  RULE SELECTION — slab-based matching on REMARKS
# # # # ─────────────────────────────────────────────────────────────────────────────

# # # def select_po(rules_list, p):
# # #     if not rules_list:
# # #         return None
# # #     for r in rules_list:
# # #         rem = str(r.get("REMARKS", "")).upper().strip()
# # #         if not rem or rem == "NIL" or rem == "ALL FUEL":
# # #             return r["PO"]
# # #         m_below = re.search(r'BELOW\s+(\d+)%', rem)
# # #         if m_below and p <= float(m_below.group(1)):
# # #             return r["PO"]
# # #         m_range = re.search(r'(\d+)%\s+TO\s+(\d+)%', rem)
# # #         if m_range:
# # #             lo, hi = float(m_range.group(1)), float(m_range.group(2))
# # #             if lo <= p <= hi:
# # #                 return r["PO"]
# # #         m_above = re.search(r'ABOVE\s+(\d+)%', rem)
# # #         if m_above and p > float(m_above.group(1)):
# # #             return r["PO"]
# # #     return rules_list[0]["PO"]


# # # # ─────────────────────────────────────────────────────────────────────────────
# # # #  INSURER MATCHING HELPER
# # # # ─────────────────────────────────────────────────────────────────────────────

# # # def insurer_matches(rule_insurer_str, ins):
# # #     """
# # #     Check if the row's insurer matches a rule's INSURER field.

# # #     Handles entries like "Tata- Comp" where the rule token contains the
# # #     insurer name plus extra words/punctuation (e.g. product suffix).

# # #     Strategy per comma-separated token in the rule:
# # #       1. Exact match after normalisation.
# # #       2. First word of token equals insurer  →  "TATA- COMP" first word = "TATA".
# # #       3. Insurer is a substring of the token as a final fallback.
# # #     """
# # #     ri = str(rule_insurer_str).strip().upper()
# # #     if ri == "ALL COMPANIES":
# # #         return True

# # #     ins_norm = ins.strip().upper()
# # #     for token in [x.strip().upper() for x in ri.split(",")]:
# # #         if ins_norm == token:
# # #             return True
# # #         # Strip punctuation so "TATA- COMP" → ["TATA", "COMP"]
# # #         token_words = re.sub(r'[^A-Z0-9 ]', ' ', token).split()
# # #         if token_words and token_words[0] == ins_norm:
# # #             return True
# # #         if ins_norm in token:
# # #             return True

# # #     return False


# # # def filter_by_insurer(rules, ins):
# # #     """
# # #     Return best matching rules for given insurer:
# # #     1. Specific match (not 'All Companies', not 'Rest of Companies')
# # #     2. 'All Companies'
# # #     3. 'Rest of Companies'
# # #     """
# # #     specific = [r for r in rules
# # #                 if str(r.get("INSURER","")).strip().upper() not in ("ALL COMPANIES","REST OF COMPANIES")
# # #                 and insurer_matches(r.get("INSURER",""), ins)]
# # #     if specific:
# # #         return specific

# # #     all_co = [r for r in rules if str(r.get("INSURER","")).strip().upper() == "ALL COMPANIES"]
# # #     if all_co:
# # #         return all_co

# # #     rest = [r for r in rules if str(r.get("INSURER","")).strip().upper() == "REST OF COMPANIES"]
# # #     return rest


# # # # ─────────────────────────────────────────────────────────────────────────────
# # # #  DETERMINE JSON SEGMENT from row data
# # # # ─────────────────────────────────────────────────────────────────────────────

# # # def get_json_lob_and_segment(sub_product_name, segment, vehicle_type_id,
# # #                               from_wt, to_wt, company_code, is_od):
# # #     """
# # #     Returns (lob, json_segment) tuple for rule lookup.
# # #     lob is the JSON LOB string.
# # #     json_segment is the JSON SEGMENT string.
# # #     Returns (None, None) if not mappable.
# # #     """
# # #     sp = str(sub_product_name).strip()
# # #     seg = str(segment).strip().upper()
# # #     vt_id = int(vehicle_type_id) if pd.notna(vehicle_type_id) else 0
# # #     ins_upper = str(company_code).strip().upper()

# # #     # ── TWO WHEELER ──────────────────────────────────────────────────────────
# # #     if sp == "Two Wheeler":
# # #         lob = "TW"
# # #         if is_od or "COMP" in seg or "SAOD" in seg or "OD" in seg:
# # #             return lob, "TW SAOD + COMP"
# # #         else:  # TP Only
# # #             return lob, "TW TP"

# # #     # ── PRIVATE CAR ──────────────────────────────────────────────────────────
# # #     if sp == "Private Car":
# # #         lob = "PVT CAR"
# # #         if is_od or "COMP" in seg or "SAOD" in seg or "OD" in seg:
# # #             return lob, "PVT CAR COMP + SAOD"
# # #         else:
# # #             return lob, "PVT CAR TP"

# # #     # ── PASSENGER VEHICLE ────────────────────────────────────────────────────
# # #     if sp == "Passenger Vehicle":
# # #         # TAXI
# # #         if vt_id in TAXI_VEHICLE_IDS:
# # #             return "TAXI", "TAXI"

# # #         # STAFF BUS
# # #         if vt_id in STAFF_BUS_IDS:
# # #             return "BUS", "STAFF BUS"

# # #         # SCHOOL BUS
# # #         if vt_id in SCHOOL_BUS_IDS:
# # #             return "BUS", "SCHOOL BUS"

# # #         # ROUTE/PASSENGER BUS → treat as STAFF BUS (general BUS)
# # #         if vt_id in ROUTE_BUS_IDS:
# # #             return "BUS", "STAFF BUS"

# # #         # 3-wheelers (Auto, Electric Rikshaw, PCV 2W etc.)
# # #         if vt_id in PCV_3W_IDS:
# # #             return "GCV, PCV 3W", "All GVW & PCV 3W, GCV 3W"

# # #         # Tempo Traveller — treat as staff bus
# # #         return "BUS", "STAFF BUS"

# # #     # ── GOODS VEHICLE ────────────────────────────────────────────────────────
# # #     if sp == "Goods Vehicle":
# # #         from_w = float(from_wt) if pd.notna(from_wt) else 0
# # #         to_w   = float(to_wt)   if pd.notna(to_wt)   else 99999

# # #         # 3-Wheeler goods
# # #         if vt_id in GCV_3W_IDS:
# # #             return "GCV, PCV 3W", "All GVW & PCV 3W, GCV 3W"

# # #         # Upto 2.5T GVW + special insurers
# # #         if from_w == 0 and to_w <= 2500 and ins_upper in SPECIAL_GCV_INSURERS:
# # #             return "GCV, PCV 3W", "Upto 2.5 GVW"

# # #         # Everything else (inc. upto 2.5T with other insurers)
# # #         return "GCV, PCV 3W", "All GVW & PCV 3W, GCV 3W"

# # #     # ── MISCELLANEOUS VEHICLE (Tractor, Harvester etc.) ───────────────────────
# # #     if sp == "Miscellaneous Vehicle":
# # #         return "MISD", "Misd, Tractor"

# # #     return None, None


# # # # ─────────────────────────────────────────────────────────────────────────────
# # # #  CORE FORMULA ENGINE
# # # # ─────────────────────────────────────────────────────────────────────────────

# # # def compute_payout(payin, sub_product_name, segment, company_code,
# # #                    rto_group_name, vehicle_type_id, from_wt, to_wt,
# # #                    rules, is_od=True):
# # #     """
# # #     Returns (payout_value, explanation_dict).
# # #     explanation_dict has keys: lob, segment, insurer_matched, po_formula,
# # #     remarks_slab, odisha_deduction, calculation_note
# # #     """
# # #     explanation = {
# # #         "lob": "",
# # #         "segment": "",
# # #         "insurer_matched": "",
# # #         "po_formula": "",
# # #         "remarks_slab": "",
# # #         "odisha_deduction": "",
# # #         "calculation_note": "",
# # #     }

# # #     try:
# # #         p = float(payin)
# # #     except (TypeError, ValueError):
# # #         explanation["calculation_note"] = "Invalid payin value"
# # #         return 0.0, explanation
# # #     if p == 0:
# # #         explanation["calculation_note"] = "Payin is 0 — no payout"
# # #         return 0.0, explanation

# # #     ins = str(company_code).strip().upper() if pd.notna(company_code) else ""
# # #     loc = str(rto_group_name).strip().upper() if pd.notna(rto_group_name) else ""

# # #     lob, json_seg = get_json_lob_and_segment(
# # #         sub_product_name, segment, vehicle_type_id, from_wt, to_wt, company_code, is_od
# # #     )

# # #     explanation["lob"]     = lob     if lob     else "NOT MAPPED"
# # #     explanation["segment"] = json_seg if json_seg else "NOT MAPPED"

# # #     if lob is None:
# # #         p_out = floor_payout(p * 0.90)
# # #         explanation["po_formula"]        = "90% of Payin (fallback)"
# # #         explanation["insurer_matched"]   = "N/A"
# # #         explanation["remarks_slab"]      = "N/A"
# # #         explanation["calculation_note"]  = f"No LOB mapping found. Fallback: floor({p} × 0.90) = {p_out}"
# # #     else:
# # #         seg_rules      = [r for r in rules if r.get("LOB") == lob and r.get("SEGMENT") == json_seg]
# # #         candidate_rules = filter_by_insurer(seg_rules, ins)
# # #         selected_po    = select_po(candidate_rules, p)

# # #         if candidate_rules:
# # #             explanation["insurer_matched"] = str(candidate_rules[0].get("INSURER", ""))
# # #             # Find which slab was picked
# # #             for r in candidate_rules:
# # #                 rem = str(r.get("REMARKS","")).upper().strip()
# # #                 if not rem or rem == "NIL" or rem == "ALL FUEL":
# # #                     explanation["remarks_slab"] = r.get("REMARKS","NIL")
# # #                     break
# # #                 m_below = re.search(r'BELOW\s+(\d+)%', rem)
# # #                 if m_below and p <= float(m_below.group(1)):
# # #                     explanation["remarks_slab"] = r.get("REMARKS","")
# # #                     break
# # #                 m_range = re.search(r'(\d+)%\s+TO\s+(\d+)%', rem)
# # #                 if m_range:
# # #                     lo, hi = float(m_range.group(1)), float(m_range.group(2))
# # #                     if lo <= p <= hi:
# # #                         explanation["remarks_slab"] = r.get("REMARKS","")
# # #                         break
# # #                 m_above = re.search(r'ABOVE\s+(\d+)%', rem)
# # #                 if m_above and p > float(m_above.group(1)):
# # #                     explanation["remarks_slab"] = r.get("REMARKS","")
# # #                     break
# # #         else:
# # #             explanation["insurer_matched"] = "No matching rule"
# # #             explanation["remarks_slab"]    = "N/A"

# # #         if selected_po is None:
# # #             p_out = floor_payout(p * 0.90)
# # #             explanation["po_formula"]       = "90% of Payin (fallback — no rule matched)"
# # #             explanation["calculation_note"] = f"Fallback: floor({p} × 0.90) = {p_out}"
# # #         else:
# # #             explanation["po_formula"] = selected_po
# # #             p_out = parse_po_to_payout(selected_po, p)
# # #             # Build human-readable calculation note
# # #             po_up = str(selected_po).strip().upper()
# # #             if re.search(r'\d+%\s*OF PAYIN', po_up):
# # #                 pct = float(re.search(r'(\d+(?:\.\d+)?)%', po_up).group(1))
# # #                 explanation["calculation_note"] = f"floor({p} × {pct}/100) = {p_out}"
# # #             elif "PAYIN + 1" in po_up:
# # #                 explanation["calculation_note"] = f"floor({p} + 1) = {p_out}"
# # #             elif "LESS 2% OF PAYIN" in po_up:
# # #                 explanation["calculation_note"] = f"floor({p} - 2) = {p_out}"
# # #             elif re.fullmatch(r'-(\d+(?:\.\d+)?)%', po_up):
# # #                 ded = float(re.search(r'-(\d+(?:\.\d+)?)%', po_up).group(1))
# # #                 explanation["calculation_note"] = f"floor({p} - {ded}) = {p_out}"
# # #             elif re.search(r'(\d+(?:\.\d+)?)%\s*PO', po_up):
# # #                 explanation["calculation_note"] = f"Fixed PO = {p_out}"
# # #             else:
# # #                 explanation["calculation_note"] = f"= {p_out}"

# # #     # ── ODISHA OVERRIDE (additional deduction on top) ─────────────────────────
# # #     if "ODISHA" in loc:
# # #         odisha_rules = [r for r in rules
# # #                         if r.get("LOCATION") == "ODISHA"
# # #                         and r.get("SEGMENT") == "ALL SEGMENT"
# # #                         and str(r.get("INSURER","")).strip().upper() == "ALL COMPANIES"]
# # #         ded_po = select_po(odisha_rules, p)
# # #         if ded_po:
# # #             ded_str = str(ded_po).strip().upper()
# # #             m = re.search(r'-(\d+(?:\.\d+)?)%', ded_str)
# # #             if m:
# # #                 ded = float(m.group(1))
# # #                 p_out = floor_payout(p - ded)   # slab on payin, deduction on payin
# # #                 explanation["odisha_deduction"] = (
# # #                     f"Odisha override ({ded_po}): floor(payin {p} - {ded}) = {p_out}"
# # #                 )

# # #     result = max(0.0, p_out)
# # #     return result, explanation


# # # # ─────────────────────────────────────────────────────────────────────────────
# # # #  MAIN
# # # # ─────────────────────────────────────────────────────────────────────────────

# # # def main():
# # #     print("\n" + "="*70)
# # #     print("  Payin-Config — Payout Recalculator")
# # #     print("="*70)

# # #     json_path  = input("\nEnter path to payout_rules.json : ").strip().strip('"')
# # #     input_path = input("Enter path to PayinConfig.xlsx  : ").strip().strip('"')
# # #     output_path= input("Enter output file path (blank=auto): ").strip().strip('"')

# # #     try:
# # #         with open(json_path) as f:
# # #             rules = json.load(f)
# # #         print(f"  Loaded {len(rules)} rules from {json_path}")
# # #     except Exception as e:
# # #         print(f"\n[ERROR] Failed to load JSON: {e}"); return

# # #     if not output_path:
# # #         base, ext = os.path.splitext(input_path)
# # #         ts = datetime.now().strftime("%Y%m%d_%H%M%S")
# # #         output_path = f"{base}_recalculated_{ts}{ext}"

# # #     print(f"\n  Reading: {input_path}")
# # #     df = pd.read_excel(input_path)
# # #     df.columns = [c.strip() for c in df.columns]
# # #     total = len(df)
# # #     print(f"  Rows   : {total}")

# # #     required = ['payin_od_rate','payin_tp_rate','payout_od_rate','payout_tp_rate',
# # #                 'sub_product_name','segment','company_code','rto_group_name',
# # #                 'vehicle_type_id','from_weightage_kg','to_weightage_kg']
# # #     missing = [c for c in required if c not in df.columns]
# # #     if missing:
# # #         print(f"\n[ERROR] Missing columns: {missing}"); return

# # #     new_od, new_tp = [], []
# # #     changed_od = changed_tp = processed_od = processed_tp = 0

# # #     # Explanation column lists
# # #     od_lob, od_seg, od_ins, od_po, od_slab, od_odisha, od_note = [], [], [], [], [], [], []
# # #     tp_lob, tp_seg, tp_ins, tp_po, tp_slab, tp_odisha, tp_note = [], [], [], [], [], [], []

# # #     for _, row in df.iterrows():
# # #         sp   = row['sub_product_name']
# # #         seg  = row['segment']
# # #         ins  = row['company_code']
# # #         loc  = row['rto_group_name']
# # #         vt   = row['vehicle_type_id']
# # #         f_wt = row['from_weightage_kg']
# # #         t_wt = row['to_weightage_kg']

# # #         # OD
# # #         payin_od = row['payin_od_rate']
# # #         old_od   = row['payout_od_rate']
# # #         if pd.notna(payin_od) and float(payin_od) != 0:
# # #             processed_od += 1
# # #             calc_od, expl_od = compute_payout(payin_od, sp, seg, ins, loc, vt, f_wt, t_wt, rules, is_od=True)
# # #             new_od.append(calc_od)
# # #             if abs(float(old_od) - calc_od) > 0.001:
# # #                 changed_od += 1
# # #         else:
# # #             calc_od = 0.0 if pd.isna(old_od) else old_od
# # #             new_od.append(calc_od)
# # #             expl_od = {"lob":"","segment":"","insurer_matched":"","po_formula":"",
# # #                        "remarks_slab":"","odisha_deduction":"","calculation_note":"Payin is 0"}
# # #         od_lob.append(expl_od["lob"])
# # #         od_seg.append(expl_od["segment"])
# # #         od_ins.append(expl_od["insurer_matched"])
# # #         od_po.append(expl_od["po_formula"])
# # #         od_slab.append(expl_od["remarks_slab"])
# # #         od_odisha.append(expl_od["odisha_deduction"])
# # #         od_note.append(expl_od["calculation_note"])

# # #         # TP
# # #         payin_tp = row['payin_tp_rate']
# # #         old_tp   = row['payout_tp_rate']
# # #         if pd.notna(payin_tp) and float(payin_tp) != 0:
# # #             processed_tp += 1
# # #             calc_tp, expl_tp = compute_payout(payin_tp, sp, seg, ins, loc, vt, f_wt, t_wt, rules, is_od=False)
# # #             new_tp.append(calc_tp)
# # #             if abs(float(old_tp) - calc_tp) > 0.001:
# # #                 changed_tp += 1
# # #         else:
# # #             calc_tp = 0.0 if pd.isna(old_tp) else old_tp
# # #             new_tp.append(calc_tp)
# # #             expl_tp = {"lob":"","segment":"","insurer_matched":"","po_formula":"",
# # #                        "remarks_slab":"","odisha_deduction":"","calculation_note":"Payin is 0"}
# # #         tp_lob.append(expl_tp["lob"])
# # #         tp_seg.append(expl_tp["segment"])
# # #         tp_ins.append(expl_tp["insurer_matched"])
# # #         tp_po.append(expl_tp["po_formula"])
# # #         tp_slab.append(expl_tp["remarks_slab"])
# # #         tp_odisha.append(expl_tp["odisha_deduction"])
# # #         tp_note.append(expl_tp["calculation_note"])

# # #     df['payout_od_rate'] = new_od
# # #     df['payout_tp_rate'] = new_tp

# # #     # Append OD explanation columns 
# # #     df['od_rule_lob']            = od_lob
# # #     df['od_rule_segment']        = od_seg
# # #     df['od_rule_insurer']        = od_ins
# # #     df['od_rule_po_formula']     = od_po
# # #     df['od_rule_slab']           = od_slab
# # #     df['od_rule_odisha']         = od_odisha
# # #     df['od_rule_calculation']    = od_note

# # #     # Append TP explanation columns
# # #     df['tp_rule_lob']            = tp_lob
# # #     df['tp_rule_segment']        = tp_seg
# # #     df['tp_rule_insurer']        = tp_ins
# # #     df['tp_rule_po_formula']     = tp_po
# # #     df['tp_rule_slab']           = tp_slab
# # #     df['tp_rule_odisha']         = tp_odisha
# # #     df['tp_rule_calculation']    = tp_note

# # #     df.to_excel(output_path, index=False)

# # #     print(f"\n{'='*70}")
# # #     print(f"  COMPLETED")
# # #     print(f"  Total rows           : {total}")
# # #     print(f"  OD rows recalculated : {processed_od}   changed: {changed_od}")
# # #     print(f"  TP rows recalculated : {processed_tp}   changed: {changed_tp}")
# # #     print(f"  Output saved to      : {output_path}")
# # #     print(f"{'='*70}")

# # #     # Sample preview
# # #     sample_cols = ['sub_product_name','segment','company_code','rto_group_name',
# # #                    'vehicle_type_id','payin_od_rate','payout_od_rate',
# # #                    'payin_tp_rate','payout_tp_rate']
# # #     sample = df[df['payin_od_rate'] > 0].head(15)[sample_cols]
# # #     print("\n  Sample output (first 15 non-zero OD rows):\n")
# # #     print(sample.to_string(index=False))
# # #     print()


# # # if __name__ == "__main__":
# # #     main()


# # """
# # Payin-Config — Payout Recalculator
# # ===================================
# # Reads PayinConfig.xlsx + master Vehicle Type table (hardcoded from provided data)
# # and recomputes payout_od_rate / payout_tp_rate using JSON payout rules.

# # Usage:
# #     python recalculate_payout.py
# # """

# # import pandas as pd
# # import math
# # import os
# # import json
# # import re
# # from datetime import datetime

# # # ─────────────────────────────────────────────────────────────────────────────
# # #  MASTER — Vehicle Type lookup (from master file Vehicle type worksheet)
# # # ─────────────────────────────────────────────────────────────────────────────

# # VEHICLE_TYPE_MASTER = {
# #     1:  {"vehicle_type": "Agriculture Tractor",          "sub_product_name": "Miscellaneous Vehicle"},
# #     2:  {"vehicle_type": "Non Tractor",                  "sub_product_name": "Miscellaneous Vehicle"},
# #     3:  {"vehicle_type": "Truck",                        "sub_product_name": "Goods Vehicle"},
# #     4:  {"vehicle_type": "Good Carring Tractor",         "sub_product_name": "Goods Vehicle"},
# #     5:  {"vehicle_type": "Tanker",                       "sub_product_name": "Goods Vehicle"},
# #     6:  {"vehicle_type": "Pickup",                       "sub_product_name": "Goods Vehicle"},
# #     7:  {"vehicle_type": "GCV 3W Delivery Van",          "sub_product_name": "Goods Vehicle"},
# #     8:  {"vehicle_type": "Taxi_CAB",                     "sub_product_name": "Passenger Vehicle"},
# #     9:  {"vehicle_type": "Electric Rikshaw",             "sub_product_name": "Passenger Vehicle"},
# #     10: {"vehicle_type": "Tempo Traveller",              "sub_product_name": "Passenger Vehicle"},
# #     11: {"vehicle_type": "School Bus",                   "sub_product_name": "Passenger Vehicle"},
# #     12: {"vehicle_type": "Passanger Bus",                "sub_product_name": "Passenger Vehicle"},
# #     13: {"vehicle_type": "Auto rikshaw",                 "sub_product_name": "Passenger Vehicle"},
# #     14: {"vehicle_type": "3W Tipper",                    "sub_product_name": "Goods Vehicle"},
# #     15: {"vehicle_type": "PCV 2W",                       "sub_product_name": "Passenger Vehicle"},
# #     16: {"vehicle_type": "GCV 2W",                       "sub_product_name": "Goods Vehicle"},
# #     17: {"vehicle_type": "TW Scooter",                   "sub_product_name": "Two Wheeler"},
# #     18: {"vehicle_type": "TW Bike",                      "sub_product_name": "Two Wheeler"},
# #     19: {"vehicle_type": "Private Car",                  "sub_product_name": "Private Car"},
# #     20: {"vehicle_type": "TW Electric Bike",             "sub_product_name": "Two Wheeler"},
# #     21: {"vehicle_type": "Electric GCV 3W Delivery Van", "sub_product_name": "Goods Vehicle"},
# #     22: {"vehicle_type": "Private Car Electric",         "sub_product_name": "Private Car"},
# #     23: {"vehicle_type": "Trailer",                      "sub_product_name": "Goods Vehicle"},
# #     24: {"vehicle_type": "Electric Pickup",              "sub_product_name": "Goods Vehicle"},
# #     25: {"vehicle_type": "Tipper",                       "sub_product_name": "Goods Vehicle"},
# #     26: {"vehicle_type": "TW Electric Scooter",          "sub_product_name": "Two Wheeler"},
# #     27: {"vehicle_type": "PC Petrol / Electric Hybrid",  "sub_product_name": "Private Car"},
# #     28: {"vehicle_type": "Staff Bus",                    "sub_product_name": "Passenger Vehicle"},
# #     29: {"vehicle_type": "Agriculture Harvester",        "sub_product_name": "Miscellaneous Vehicle"},
# #     30: {"vehicle_type": "Electric PCV 2W",              "sub_product_name": "Passenger Vehicle"},
# #     31: {"vehicle_type": "Electric Taxi_CAB",            "sub_product_name": "Passenger Vehicle"},
# #     32: {"vehicle_type": "Route Bus",                    "sub_product_name": "Passenger Vehicle"},
# # }

# # # IDs considered as TAXI
# # TAXI_VEHICLE_IDS = {8, 31}  # Taxi_CAB, Electric Taxi_CAB

# # # IDs considered as STAFF BUS
# # STAFF_BUS_IDS = {28}  # Staff Bus

# # # IDs considered as SCHOOL BUS
# # SCHOOL_BUS_IDS = {11}  # School Bus

# # # IDs considered as BUS (any bus — route/passenger)
# # ROUTE_BUS_IDS = {12, 32}  # Passanger Bus, Route Bus

# # # IDs considered as GCV 3-Wheeler goods
# # GCV_3W_IDS = {7, 14, 21}  # GCV 3W Delivery Van, 3W Tipper, Electric GCV 3W

# # # IDs considered as Passenger 3-Wheeler (auto etc.)
# # PCV_3W_IDS = {9, 13, 15, 30}  # Electric Rikshaw, Auto rikshaw, PCV 2W, Electric PCV 2W

# # # Insurers to match for Upto 2.5 GVW special rule
# # SPECIAL_GCV_INSURERS = {"RELIANCE", "SBI"}

# # # =============================================================================
# # #  RTO MASTER LOOKUP — maps city / code / group name / state -> state name
# # # =============================================================================

# # def build_rto_state_lookup(master_path):
# #     """
# #     Reads the 'RTO Master' sheet and builds a dict:
# #         uppercased_key -> state_name_uppercase

# #     Keys added per row:
# #       - rto_code  (e.g. "OD01")
# #       - rto_name  (e.g. "Bhubaneswar")
# #       - state_name (maps to itself, so a direct state hit also works)
# #     """
# #     try:
# #         rto_df = pd.read_excel(master_path, sheet_name="RTO Master")
# #         rto_df.columns = [c.strip() for c in rto_df.columns]
# #     except Exception as e:
# #         print(f"  [WARN] Could not load RTO Master sheet: {e}")
# #         print(f"  [WARN] Odisha detection will fall back to substring match only.")
# #         return {}

# #     lookup = {}
# #     for _, row in rto_df.iterrows():
# #         state = str(row.get("state_name", "")).strip()
# #         if not state or state.lower() == "nan":
# #             continue
# #         state_up = state.upper()

# #         for col in ["rto_code", "rto_name"]:
# #             val = str(row.get(col, "")).strip().upper()
# #             if val and val != "NAN":
# #                 lookup[val] = state_up

# #         # state maps to itself
# #         lookup[state_up] = state_up

# #     print(f"  RTO Master loaded: {len(lookup)} lookup entries.")
# #     return lookup


# # def resolve_state(rto_group_name, rto_lookup):
# #     """
# #     Map any rto_group_name value to its state name (uppercase).

# #     Tries in order:
# #       1. Exact match (uppercased)
# #       2. Any lookup KEY found inside the rto_group_name string
# #          e.g. "Western Odisha" contains "ODISHA" -> "ODISHA"
# #          e.g. "BHUBANESWAR GROUP" contains "BHUBANESWAR" -> "ODISHA"
# #       3. rto_group_name found inside any lookup key (loose reverse match)
# #     Returns the resolved state or the original value uppercased if no match.
# #     """
# #     if not rto_lookup:
# #         return str(rto_group_name).strip().upper()

# #     val = str(rto_group_name).strip().upper()

# #     # 1. Exact
# #     if val in rto_lookup:
# #         return rto_lookup[val]

# #     # 2. Any key is a substring of val  (e.g. val="WESTERN ODISHA", key="ODISHA")
# #     for key, state in rto_lookup.items():
# #         if key in val:
# #             return state

# #     # 3. val is a substring of any key
# #     for key, state in rto_lookup.items():
# #         if val in key:
# #             return state

# #     return val


# # # ─────────────────────────────────────────────────────────────────────────────
# # #  FLOOR
# # # ─────────────────────────────────────────────────────────────────────────────

# # def floor_payout(value):
# #     return float(math.floor(float(value)))


# # # ─────────────────────────────────────────────────────────────────────────────
# # #  PO STRING PARSER
# # # ─────────────────────────────────────────────────────────────────────────────

# # def parse_po_to_payout(po_str, p):
# #     po_str = str(po_str).strip().upper()

# #     if re.search(r'\d+%\s*OF PAYIN', po_str):
# #         percent = float(re.search(r'(\d+(?:\.\d+)?)%', po_str).group(1))
# #         return floor_payout(p * (percent / 100))

# #     if "PAYIN + 1" in po_str:
# #         return floor_payout(p + 1)

# #     if "LESS 2% OF PAYIN" in po_str:
# #         return floor_payout(p - 2)

# #     # "-3%", "-4%", "-5%"
# #     m = re.fullmatch(r'-(\d+(?:\.\d+)?)%', po_str)
# #     if m:
# #         ded = float(m.group(1))
# #         return floor_payout(p - ded)

# #     # "21% PO" — fixed payout
# #     m = re.search(r'(\d+(?:\.\d+)?)%\s*PO', po_str)
# #     if m:
# #         return floor_payout(float(m.group(1)))

# #     # Fallback
# #     return floor_payout(p * 0.90)


# # # ─────────────────────────────────────────────────────────────────────────────
# # #  RULE SELECTION — slab-based matching on REMARKS
# # # ─────────────────────────────────────────────────────────────────────────────

# # def select_po(rules_list, p):
# #     if not rules_list:
# #         return None
# #     for r in rules_list:
# #         rem = str(r.get("REMARKS", "")).upper().strip()
# #         if not rem or rem == "NIL" or rem == "ALL FUEL":
# #             return r["PO"]
# #         m_below = re.search(r'BELOW\s+(\d+)%', rem)
# #         if m_below and p <= float(m_below.group(1)):
# #             return r["PO"]
# #         m_range = re.search(r'(\d+)%\s+TO\s+(\d+)%', rem)
# #         if m_range:
# #             lo, hi = float(m_range.group(1)), float(m_range.group(2))
# #             if lo <= p <= hi:
# #                 return r["PO"]
# #         m_above = re.search(r'ABOVE\s+(\d+)%', rem)
# #         if m_above and p > float(m_above.group(1)):
# #             return r["PO"]
# #     return rules_list[0]["PO"]


# # # ─────────────────────────────────────────────────────────────────────────────
# # #  INSURER MATCHING HELPER
# # # ─────────────────────────────────────────────────────────────────────────────

# # def insurer_matches(rule_insurer_str, ins):
# #     """
# #     Check if the row's insurer matches a rule's INSURER field.

# #     Handles entries like "Tata- Comp" where the rule token contains the
# #     insurer name plus extra words/punctuation (e.g. product suffix).

# #     Strategy per comma-separated token in the rule:
# #       1. Exact match after normalisation.
# #       2. First word of token equals insurer  →  "TATA- COMP" first word = "TATA".
# #       3. Insurer is a substring of the token as a final fallback.
# #     """
# #     ri = str(rule_insurer_str).strip().upper()
# #     if ri == "ALL COMPANIES":
# #         return True

# #     ins_norm = ins.strip().upper()
# #     for token in [x.strip().upper() for x in ri.split(",")]:
# #         if ins_norm == token:
# #             return True
# #         # Strip punctuation so "TATA- COMP" → ["TATA", "COMP"]
# #         token_words = re.sub(r'[^A-Z0-9 ]', ' ', token).split()
# #         if token_words and token_words[0] == ins_norm:
# #             return True
# #         if ins_norm in token:
# #             return True

# #     return False


# # def filter_by_insurer(rules, ins):
# #     """
# #     Return best matching rules for given insurer:
# #     1. Specific match (not 'All Companies', not 'Rest of Companies')
# #     2. 'All Companies'
# #     3. 'Rest of Companies'
# #     """
# #     specific = [r for r in rules
# #                 if str(r.get("INSURER","")).strip().upper() not in ("ALL COMPANIES","REST OF COMPANIES")
# #                 and insurer_matches(r.get("INSURER",""), ins)]
# #     if specific:
# #         return specific

# #     all_co = [r for r in rules if str(r.get("INSURER","")).strip().upper() == "ALL COMPANIES"]
# #     if all_co:
# #         return all_co

# #     rest = [r for r in rules if str(r.get("INSURER","")).strip().upper() == "REST OF COMPANIES"]
# #     return rest


# # # ─────────────────────────────────────────────────────────────────────────────
# # #  DETERMINE JSON SEGMENT from row data
# # # ─────────────────────────────────────────────────────────────────────────────

# # def get_json_lob_and_segment(sub_product_name, segment, vehicle_type_id,
# #                               from_wt, to_wt, company_code, is_od):
# #     """
# #     Returns (lob, json_segment) tuple for rule lookup.
# #     lob is the JSON LOB string.
# #     json_segment is the JSON SEGMENT string.
# #     Returns (None, None) if not mappable.
# #     """
# #     sp = str(sub_product_name).strip()
# #     seg = str(segment).strip().upper()
# #     vt_id = int(vehicle_type_id) if pd.notna(vehicle_type_id) else 0
# #     ins_upper = str(company_code).strip().upper()

# #     # ── TWO WHEELER ──────────────────────────────────────────────────────────
# #     if sp == "Two Wheeler":
# #         lob = "TW"
# #         if is_od or "COMP" in seg or "SAOD" in seg or "OD" in seg:
# #             return lob, "TW SAOD + COMP"
# #         else:  # TP Only
# #             return lob, "TW TP"

# #     # ── PRIVATE CAR ──────────────────────────────────────────────────────────
# #     if sp == "Private Car":
# #         lob = "PVT CAR"
# #         if is_od or "COMP" in seg or "SAOD" in seg or "OD" in seg:
# #             return lob, "PVT CAR COMP + SAOD"
# #         else:
# #             return lob, "PVT CAR TP"

# #     # ── PASSENGER VEHICLE ────────────────────────────────────────────────────
# #     if sp == "Passenger Vehicle":
# #         # TAXI
# #         if vt_id in TAXI_VEHICLE_IDS:
# #             return "TAXI", "TAXI"

# #         # STAFF BUS
# #         if vt_id in STAFF_BUS_IDS:
# #             return "BUS", "STAFF BUS"

# #         # SCHOOL BUS
# #         if vt_id in SCHOOL_BUS_IDS:
# #             return "BUS", "SCHOOL BUS"

# #         # ROUTE/PASSENGER BUS → treat as STAFF BUS (general BUS)
# #         if vt_id in ROUTE_BUS_IDS:
# #             return "BUS", "STAFF BUS"

# #         # 3-wheelers (Auto, Electric Rikshaw, PCV 2W etc.)
# #         if vt_id in PCV_3W_IDS:
# #             return "GCV, PCV 3W", "All GVW & PCV 3W, GCV 3W"

# #         # Tempo Traveller — treat as staff bus
# #         return "BUS", "STAFF BUS"

# #     # ── GOODS VEHICLE ────────────────────────────────────────────────────────
# #     if sp == "Goods Vehicle":
# #         from_w = float(from_wt) if pd.notna(from_wt) else 0
# #         to_w   = float(to_wt)   if pd.notna(to_wt)   else 99999

# #         # 3-Wheeler goods
# #         if vt_id in GCV_3W_IDS:
# #             return "GCV, PCV 3W", "All GVW & PCV 3W, GCV 3W"

# #         # Upto 2.5T GVW + special insurers
# #         if from_w == 0 and to_w <= 2500 and ins_upper in SPECIAL_GCV_INSURERS:
# #             return "GCV, PCV 3W", "Upto 2.5 GVW"

# #         # Everything else (inc. upto 2.5T with other insurers)
# #         return "GCV, PCV 3W", "All GVW & PCV 3W, GCV 3W"

# #     # ── MISCELLANEOUS VEHICLE (Tractor, Harvester etc.) ───────────────────────
# #     if sp == "Miscellaneous Vehicle":
# #         return "MISD", "Misd, Tractor"

# #     return None, None


# # # ─────────────────────────────────────────────────────────────────────────────
# # #  CORE FORMULA ENGINE
# # # ─────────────────────────────────────────────────────────────────────────────

# # def compute_payout(payin, sub_product_name, segment, company_code,
# #                    rto_group_name, vehicle_type_id, from_wt, to_wt,
# #                    rules, is_od=True, rto_lookup=None):
# #     """
# #     Returns (payout_value, explanation_dict).
# #     explanation_dict has keys: lob, segment, insurer_matched, po_formula,
# #     remarks_slab, odisha_deduction, calculation_note
# #     """
# #     explanation = {
# #         "lob": "",
# #         "segment": "",
# #         "insurer_matched": "",
# #         "po_formula": "",
# #         "remarks_slab": "",
# #         "odisha_deduction": "",
# #         "calculation_note": "",
# #     }

# #     try:
# #         p = float(payin)
# #     except (TypeError, ValueError):
# #         explanation["calculation_note"] = "Invalid payin value"
# #         return 0.0, explanation
# #     if p == 0:
# #         explanation["calculation_note"] = "Payin is 0 — no payout"
# #         return 0.0, explanation

# #     ins = str(company_code).strip().upper() if pd.notna(company_code) else ""
# #     raw_loc = str(rto_group_name).strip() if pd.notna(rto_group_name) else ""
# #     loc = resolve_state(raw_loc, rto_lookup or {})

# #     lob, json_seg = get_json_lob_and_segment(
# #         sub_product_name, segment, vehicle_type_id, from_wt, to_wt, company_code, is_od
# #     )

# #     explanation["lob"]     = lob     if lob     else "NOT MAPPED"
# #     explanation["segment"] = json_seg if json_seg else "NOT MAPPED"

# #     if lob is None:
# #         p_out = floor_payout(p * 0.90)
# #         explanation["po_formula"]        = "90% of Payin (fallback)"
# #         explanation["insurer_matched"]   = "N/A"
# #         explanation["remarks_slab"]      = "N/A"
# #         explanation["calculation_note"]  = f"No LOB mapping found. Fallback: floor({p} × 0.90) = {p_out}"
# #     else:
# #         seg_rules      = [r for r in rules if r.get("LOB") == lob and r.get("SEGMENT") == json_seg]
# #         candidate_rules = filter_by_insurer(seg_rules, ins)
# #         selected_po    = select_po(candidate_rules, p)

# #         if candidate_rules:
# #             explanation["insurer_matched"] = str(candidate_rules[0].get("INSURER", ""))
# #             # Find which slab was picked
# #             for r in candidate_rules:
# #                 rem = str(r.get("REMARKS","")).upper().strip()
# #                 if not rem or rem == "NIL" or rem == "ALL FUEL":
# #                     explanation["remarks_slab"] = r.get("REMARKS","NIL")
# #                     break
# #                 m_below = re.search(r'BELOW\s+(\d+)%', rem)
# #                 if m_below and p <= float(m_below.group(1)):
# #                     explanation["remarks_slab"] = r.get("REMARKS","")
# #                     break
# #                 m_range = re.search(r'(\d+)%\s+TO\s+(\d+)%', rem)
# #                 if m_range:
# #                     lo, hi = float(m_range.group(1)), float(m_range.group(2))
# #                     if lo <= p <= hi:
# #                         explanation["remarks_slab"] = r.get("REMARKS","")
# #                         break
# #                 m_above = re.search(r'ABOVE\s+(\d+)%', rem)
# #                 if m_above and p > float(m_above.group(1)):
# #                     explanation["remarks_slab"] = r.get("REMARKS","")
# #                     break
# #         else:
# #             explanation["insurer_matched"] = "No matching rule"
# #             explanation["remarks_slab"]    = "N/A"

# #         if selected_po is None:
# #             p_out = floor_payout(p * 0.90)
# #             explanation["po_formula"]       = "90% of Payin (fallback — no rule matched)"
# #             explanation["calculation_note"] = f"Fallback: floor({p} × 0.90) = {p_out}"
# #         else:
# #             explanation["po_formula"] = selected_po
# #             p_out = parse_po_to_payout(selected_po, p)
# #             # Build human-readable calculation note
# #             po_up = str(selected_po).strip().upper()
# #             if re.search(r'\d+%\s*OF PAYIN', po_up):
# #                 pct = float(re.search(r'(\d+(?:\.\d+)?)%', po_up).group(1))
# #                 explanation["calculation_note"] = f"floor({p} × {pct}/100) = {p_out}"
# #             elif "PAYIN + 1" in po_up:
# #                 explanation["calculation_note"] = f"floor({p} + 1) = {p_out}"
# #             elif "LESS 2% OF PAYIN" in po_up:
# #                 explanation["calculation_note"] = f"floor({p} - 2) = {p_out}"
# #             elif re.fullmatch(r'-(\d+(?:\.\d+)?)%', po_up):
# #                 ded = float(re.search(r'-(\d+(?:\.\d+)?)%', po_up).group(1))
# #                 explanation["calculation_note"] = f"floor({p} - {ded}) = {p_out}"
# #             elif re.search(r'(\d+(?:\.\d+)?)%\s*PO', po_up):
# #                 explanation["calculation_note"] = f"Fixed PO = {p_out}"
# #             else:
# #                 explanation["calculation_note"] = f"= {p_out}"

# #     # ── ODISHA OVERRIDE (additional deduction on top) ─────────────────────────
# #     if loc == "ODISHA":
# #         odisha_rules = [r for r in rules
# #                         if r.get("LOCATION") == "ODISHA"
# #                         and r.get("SEGMENT") == "ALL SEGMENT"
# #                         and str(r.get("INSURER","")).strip().upper() == "ALL COMPANIES"]
# #         ded_po = select_po(odisha_rules, p)
# #         if ded_po:
# #             ded_str = str(ded_po).strip().upper()
# #             m = re.search(r'-(\d+(?:\.\d+)?)%', ded_str)
# #             if m:
# #                 ded = float(m.group(1))
# #                 p_out = floor_payout(p - ded)   # slab on payin, deduction on payin
# #                 explanation["odisha_deduction"] = (
# #                     f"Odisha override ({ded_po}): floor(payin {p} - {ded}) = {p_out}"
# #                 )

# #     result = max(0.0, p_out)
# #     return result, explanation


# # # ─────────────────────────────────────────────────────────────────────────────
# # #  MAIN
# # # ─────────────────────────────────────────────────────────────────────────────

# # def main():
# #     print("\n" + "="*70)
# #     print("  Payin-Config — Payout Recalculator")
# #     print("="*70)

# #     json_path   = input("\nEnter path to payout_rules.json : ").strip().strip('"')
# #     input_path  = input("Enter path to PayinConfig.xlsx  : ").strip().strip('"')
# #     master_path = input("Enter path to Master.xlsx        : ").strip().strip('"')
# #     output_path = input("Enter output file path (blank=auto): ").strip().strip('"')

# #     try:
# #         with open(json_path) as f:
# #             rules = json.load(f)
# #         print(f"  Loaded {len(rules)} rules from {json_path}")
# #     except Exception as e:
# #         print(f"\n[ERROR] Failed to load JSON: {e}"); return

# #     rto_lookup = build_rto_state_lookup(master_path)

# #     if not output_path:
# #         base, ext = os.path.splitext(input_path)
# #         ts = datetime.now().strftime("%Y%m%d_%H%M%S")
# #         output_path = f"{base}_recalculated_{ts}{ext}"

# #     print(f"\n  Reading: {input_path}")
# #     df = pd.read_excel(input_path)
# #     df.columns = [c.strip() for c in df.columns]
# #     total = len(df)
# #     print(f"  Rows   : {total}")

# #     required = ['payin_od_rate','payin_tp_rate','payout_od_rate','payout_tp_rate',
# #                 'sub_product_name','segment','company_code','rto_group_name',
# #                 'vehicle_type_id','from_weightage_kg','to_weightage_kg']
# #     missing = [c for c in required if c not in df.columns]
# #     if missing:
# #         print(f"\n[ERROR] Missing columns: {missing}"); return

# #     new_od, new_tp = [], []
# #     changed_od = changed_tp = processed_od = processed_tp = 0

# #     # Explanation column lists
# #     od_lob, od_seg, od_ins, od_po, od_slab, od_odisha, od_note = [], [], [], [], [], [], []
# #     tp_lob, tp_seg, tp_ins, tp_po, tp_slab, tp_odisha, tp_note = [], [], [], [], [], [], []

# #     for _, row in df.iterrows():
# #         sp   = row['sub_product_name']
# #         seg  = row['segment']
# #         ins  = row['company_code']
# #         loc  = row['rto_group_name']
# #         vt   = row['vehicle_type_id']
# #         f_wt = row['from_weightage_kg']
# #         t_wt = row['to_weightage_kg']

# #         # OD
# #         payin_od = row['payin_od_rate']
# #         old_od   = row['payout_od_rate']
# #         if pd.notna(payin_od) and float(payin_od) != 0:
# #             processed_od += 1
# #             calc_od, expl_od = compute_payout(payin_od, sp, seg, ins, loc, vt, f_wt, t_wt, rules, is_od=True, rto_lookup=rto_lookup)
# #             new_od.append(calc_od)
# #             if abs(float(old_od) - calc_od) > 0.001:
# #                 changed_od += 1
# #         else:
# #             calc_od = 0.0 if pd.isna(old_od) else old_od
# #             new_od.append(calc_od)
# #             expl_od = {"lob":"","segment":"","insurer_matched":"","po_formula":"",
# #                        "remarks_slab":"","odisha_deduction":"","calculation_note":"Payin is 0"}
# #         od_lob.append(expl_od["lob"])
# #         od_seg.append(expl_od["segment"])
# #         od_ins.append(expl_od["insurer_matched"])
# #         od_po.append(expl_od["po_formula"])
# #         od_slab.append(expl_od["remarks_slab"])
# #         od_odisha.append(expl_od["odisha_deduction"])
# #         od_note.append(expl_od["calculation_note"])

# #         # TP
# #         payin_tp = row['payin_tp_rate']
# #         old_tp   = row['payout_tp_rate']
# #         if pd.notna(payin_tp) and float(payin_tp) != 0:
# #             processed_tp += 1
# #             calc_tp, expl_tp = compute_payout(payin_tp, sp, seg, ins, loc, vt, f_wt, t_wt, rules, is_od=False, rto_lookup=rto_lookup)
# #             new_tp.append(calc_tp)
# #             if abs(float(old_tp) - calc_tp) > 0.001:
# #                 changed_tp += 1
# #         else:
# #             calc_tp = 0.0 if pd.isna(old_tp) else old_tp
# #             new_tp.append(calc_tp)
# #             expl_tp = {"lob":"","segment":"","insurer_matched":"","po_formula":"",
# #                        "remarks_slab":"","odisha_deduction":"","calculation_note":"Payin is 0"}
# #         tp_lob.append(expl_tp["lob"])
# #         tp_seg.append(expl_tp["segment"])
# #         tp_ins.append(expl_tp["insurer_matched"])
# #         tp_po.append(expl_tp["po_formula"])
# #         tp_slab.append(expl_tp["remarks_slab"])
# #         tp_odisha.append(expl_tp["odisha_deduction"])
# #         tp_note.append(expl_tp["calculation_note"])

# #     df['payout_od_rate'] = new_od
# #     df['payout_tp_rate'] = new_tp

# #     # Append OD explanation columns
# #     df['od_rule_lob']            = od_lob
# #     df['od_rule_segment']        = od_seg
# #     df['od_rule_insurer']        = od_ins
# #     df['od_rule_po_formula']     = od_po
# #     df['od_rule_slab']           = od_slab
# #     df['od_rule_odisha']         = od_odisha
# #     df['od_rule_calculation']    = od_note

# #     # Append TP explanation columns
# #     df['tp_rule_lob']            = tp_lob
# #     df['tp_rule_segment']        = tp_seg
# #     df['tp_rule_insurer']        = tp_ins
# #     df['tp_rule_po_formula']     = tp_po
# #     df['tp_rule_slab']           = tp_slab
# #     df['tp_rule_odisha']         = tp_odisha
# #     df['tp_rule_calculation']    = tp_note

# #     df.to_excel(output_path, index=False)

# #     print(f"\n{'='*70}")
# #     print(f"  COMPLETED")
# #     print(f"  Total rows           : {total}")
# #     print(f"  OD rows recalculated : {processed_od}   changed: {changed_od}")
# #     print(f"  TP rows recalculated : {processed_tp}   changed: {changed_tp}")
# #     print(f"  Output saved to      : {output_path}")
# #     print(f"{'='*70}")

# #     # Sample preview
# #     sample_cols = ['sub_product_name','segment','company_code','rto_group_name',
# #                    'vehicle_type_id','payin_od_rate','payout_od_rate',
# #                    'payin_tp_rate','payout_tp_rate']
# #     sample = df[df['payin_od_rate'] > 0].head(15)[sample_cols]
# #     print("\n  Sample output (first 15 non-zero OD rows):\n")
# #     print(sample.to_string(index=False))
# #     print()


# # if __name__ == "__main__":
# #     main()

# """
# Payin-Config — Payout Recalculator
# ===================================
# Reads PayinConfig.xlsx + master Vehicle Type table (hardcoded from provided data)
# and recomputes payout_od_rate / payout_tp_rate using JSON payout rules.

# Usage:
#     python recalculate_payout.py
# """

# import pandas as pd
# import math
# import os
# import json
# import re
# import unicodedata
# from datetime import datetime

# # ─────────────────────────────────────────────────────────────────────────────
# #  MASTER — Vehicle Type lookup (from master file Vehicle type worksheet)
# # ─────────────────────────────────────────────────────────────────────────────

# VEHICLE_TYPE_MASTER = {
#     1:  {"vehicle_type": "Agriculture Tractor",          "sub_product_name": "Miscellaneous Vehicle"},
#     2:  {"vehicle_type": "Non Tractor",                  "sub_product_name": "Miscellaneous Vehicle"},
#     3:  {"vehicle_type": "Truck",                        "sub_product_name": "Goods Vehicle"},
#     4:  {"vehicle_type": "Good Carring Tractor",         "sub_product_name": "Goods Vehicle"},
#     5:  {"vehicle_type": "Tanker",                       "sub_product_name": "Goods Vehicle"},
#     6:  {"vehicle_type": "Pickup",                       "sub_product_name": "Goods Vehicle"},
#     7:  {"vehicle_type": "GCV 3W Delivery Van",          "sub_product_name": "Goods Vehicle"},
#     8:  {"vehicle_type": "Taxi_CAB",                     "sub_product_name": "Passenger Vehicle"},
#     9:  {"vehicle_type": "Electric Rikshaw",             "sub_product_name": "Passenger Vehicle"},
#     10: {"vehicle_type": "Tempo Traveller",              "sub_product_name": "Passenger Vehicle"},
#     11: {"vehicle_type": "School Bus",                   "sub_product_name": "Passenger Vehicle"},
#     12: {"vehicle_type": "Passanger Bus",                "sub_product_name": "Passenger Vehicle"},
#     13: {"vehicle_type": "Auto rikshaw",                 "sub_product_name": "Passenger Vehicle"},
#     14: {"vehicle_type": "3W Tipper",                    "sub_product_name": "Goods Vehicle"},
#     15: {"vehicle_type": "PCV 2W",                       "sub_product_name": "Passenger Vehicle"},
#     16: {"vehicle_type": "GCV 2W",                       "sub_product_name": "Goods Vehicle"},
#     17: {"vehicle_type": "TW Scooter",                   "sub_product_name": "Two Wheeler"},
#     18: {"vehicle_type": "TW Bike",                      "sub_product_name": "Two Wheeler"},
#     19: {"vehicle_type": "Private Car",                  "sub_product_name": "Private Car"},
#     20: {"vehicle_type": "TW Electric Bike",             "sub_product_name": "Two Wheeler"},
#     21: {"vehicle_type": "Electric GCV 3W Delivery Van", "sub_product_name": "Goods Vehicle"},
#     22: {"vehicle_type": "Private Car Electric",         "sub_product_name": "Private Car"},
#     23: {"vehicle_type": "Trailer",                      "sub_product_name": "Goods Vehicle"},
#     24: {"vehicle_type": "Electric Pickup",              "sub_product_name": "Goods Vehicle"},
#     25: {"vehicle_type": "Tipper",                       "sub_product_name": "Goods Vehicle"},
#     26: {"vehicle_type": "TW Electric Scooter",          "sub_product_name": "Two Wheeler"},
#     27: {"vehicle_type": "PC Petrol / Electric Hybrid",  "sub_product_name": "Private Car"},
#     28: {"vehicle_type": "Staff Bus",                    "sub_product_name": "Passenger Vehicle"},
#     29: {"vehicle_type": "Agriculture Harvester",        "sub_product_name": "Miscellaneous Vehicle"},
#     30: {"vehicle_type": "Electric PCV 2W",              "sub_product_name": "Passenger Vehicle"},
#     31: {"vehicle_type": "Electric Taxi_CAB",            "sub_product_name": "Passenger Vehicle"},
#     32: {"vehicle_type": "Route Bus",                    "sub_product_name": "Passenger Vehicle"},
# }

# # IDs considered as TAXI
# TAXI_VEHICLE_IDS = {8, 31}  # Taxi_CAB, Electric Taxi_CAB

# # IDs considered as STAFF BUS
# STAFF_BUS_IDS = {28}  # Staff Bus

# # IDs considered as SCHOOL BUS
# SCHOOL_BUS_IDS = {11}  # School Bus

# # IDs considered as BUS (any bus — route/passenger)
# ROUTE_BUS_IDS = {12, 32}  # Passanger Bus, Route Bus

# # IDs considered as GCV 3-Wheeler goods
# GCV_3W_IDS = {7, 14, 21}  # GCV 3W Delivery Van, 3W Tipper, Electric GCV 3W

# # IDs considered as Passenger 3-Wheeler (auto etc.)
# PCV_3W_IDS = {9, 13, 15, 30}  # Electric Rikshaw, Auto rikshaw, PCV 2W, Electric PCV 2W

# # Insurers to match for Upto 2.5 GVW special rule
# SPECIAL_GCV_INSURERS = {"RELIANCE", "SBI"}

# # =============================================================================
# #  RTO MASTER LOOKUP — maps city / code / group name / state -> state name
# # =============================================================================

# def build_rto_state_lookup(master_path):
#     """
#     Reads the 'RTO Master' sheet and builds a dict:
#         uppercased_key -> state_name_uppercase

#     Keys added per row:
#       - rto_code  (e.g. "OD01")
#       - rto_name  (e.g. "Bhubaneswar")
#       - state_name (maps to itself, so a direct state hit also works)
#     """
#     try:
#         rto_df = pd.read_excel(master_path, sheet_name="RTO Master")
#         rto_df.columns = [c.strip() for c in rto_df.columns]
#     except Exception as e:
#         print(f"  [WARN] Could not load RTO Master sheet: {e}")
#         print(f"  [WARN] Odisha detection will fall back to substring match only.")
#         return {}

#     lookup = {}
#     for _, row in rto_df.iterrows():
#         state = str(row.get("state_name", "")).strip()
#         if not state or state.lower() == "nan":
#             continue
#         state_up = state.upper()

#         for col in ["rto_code", "rto_name"]:
#             val = str(row.get(col, "")).strip().upper()
#             if val and val != "NAN":
#                 lookup[val] = state_up

#         # state maps to itself
#         lookup[state_up] = state_up

#     print(f"  RTO Master loaded: {len(lookup)} lookup entries.")
#     return lookup


# def _normalize(text):
#     """Lowercase, strip accents, collapse spaces — for fuzzy comparison."""
#     text = unicodedata.normalize('NFKD', str(text)).encode('ascii', 'ignore').decode()
#     return re.sub(r'\s+', ' ', text).strip().lower()


# def _levenshtein(a, b):
#     """Classic Levenshtein edit distance between two strings."""
#     if a == b:
#         return 0
#     if not a:
#         return len(b)
#     if not b:
#         return len(a)
#     # Only keep two rows to save memory
#     prev = list(range(len(b) + 1))
#     curr = [0] * (len(b) + 1)
#     for i, ca in enumerate(a, 1):
#         curr[0] = i
#         for j, cb in enumerate(b, 1):
#             curr[j] = min(
#                 prev[j] + 1,        # deletion
#                 curr[j - 1] + 1,    # insertion
#                 prev[j - 1] + (0 if ca == cb else 1)  # substitution
#             )
#         prev, curr = curr, prev
#     return prev[len(b)]


# def _similarity(a, b):
#     """
#     Return similarity score 0-100 between two strings.
#     100 = identical, 0 = completely different.
#     """
#     a, b = _normalize(a), _normalize(b)
#     max_len = max(len(a), len(b), 1)
#     dist = _levenshtein(a, b)
#     return round((1 - dist / max_len) * 100)


# def resolve_state(rto_group_name, rto_lookup, fuzzy_threshold=75):
#     """
#     Map any rto_group_name value to its state name (uppercase).

#     Tries in order:
#       1. Exact match (uppercased)
#       2. Any lookup KEY found as substring inside the value
#          e.g. "Western Odisha" contains "ODISHA" -> "ODISHA"
#       3. Value found as substring inside any lookup key
#       4. Fuzzy match using Levenshtein similarity >= fuzzy_threshold
#          e.g. "BHUBANESHWAR" fuzzy-matches "BHUBANESWAR" -> "ODISHA"
#     Returns resolved state (uppercase) or original value uppercased if no match.
#     """
#     if not rto_lookup:
#         return str(rto_group_name).strip().upper()

#     val = str(rto_group_name).strip().upper()

#     # 1. Exact match
#     if val in rto_lookup:
#         return rto_lookup[val]

#     # 2. Any key is substring of val  e.g. val="WESTERN ODISHA", key="ODISHA"
#     for key, state in rto_lookup.items():
#         if key in val:
#             return state

#     # 3. val is substring of any key
#     for key, state in rto_lookup.items():
#         if val in key:
#             return state

#     # 4. Fuzzy match — find best scoring key above threshold
#     # Skip fuzzy for very short values (<=4 chars) — likely codes/abbreviations
#     # that already failed exact match; fuzzy would give false positives
#     if len(val) <= 4:
#         return val

#     best_score = 0
#     best_state = None
#     for key, state in rto_lookup.items():
#         # Also skip fuzzy against very short keys to avoid false matches
#         if len(key) <= 3:
#             continue
#         score = _similarity(val, key)
#         if score > best_score:
#             best_score = score
#             best_state = state

#     if best_score >= fuzzy_threshold:
#         return best_state

#     # No match found
#     return val


# # ─────────────────────────────────────────────────────────────────────────────
# #  FLOOR
# # ─────────────────────────────────────────────────────────────────────────────

# def floor_payout(value):
#     return float(math.floor(float(value)))


# # ─────────────────────────────────────────────────────────────────────────────
# #  PO STRING PARSER
# # ─────────────────────────────────────────────────────────────────────────────

# def parse_po_to_payout(po_str, p):
#     po_str = str(po_str).strip().upper()

#     if re.search(r'\d+%\s*OF PAYIN', po_str):
#         percent = float(re.search(r'(\d+(?:\.\d+)?)%', po_str).group(1))
#         return floor_payout(p * (percent / 100))

#     if "PAYIN + 1" in po_str:
#         return floor_payout(p + 1)

#     if "LESS 2% OF PAYIN" in po_str:
#         return floor_payout(p - 2)

#     # "-3%", "-4%", "-5%"
#     m = re.fullmatch(r'-(\d+(?:\.\d+)?)%', po_str)
#     if m:
#         ded = float(m.group(1))
#         return floor_payout(p - ded)

#     # "21% PO" — fixed payout
#     m = re.search(r'(\d+(?:\.\d+)?)%\s*PO', po_str)
#     if m:
#         return floor_payout(float(m.group(1)))

#     # Fallback
#     return floor_payout(p * 0.90)


# # ─────────────────────────────────────────────────────────────────────────────
# #  RULE SELECTION — slab-based matching on REMARKS
# # ─────────────────────────────────────────────────────────────────────────────

# def select_po(rules_list, p):
#     if not rules_list:
#         return None
#     for r in rules_list:
#         rem = str(r.get("REMARKS", "")).upper().strip()
#         if not rem or rem == "NIL" or rem == "ALL FUEL":
#             return r["PO"]
#         m_below = re.search(r'BELOW\s+(\d+)%', rem)
#         if m_below and p <= float(m_below.group(1)):
#             return r["PO"]
#         m_range = re.search(r'(\d+)%\s+TO\s+(\d+)%', rem)
#         if m_range:
#             lo, hi = float(m_range.group(1)), float(m_range.group(2))
#             if lo <= p <= hi:
#                 return r["PO"]
#         m_above = re.search(r'ABOVE\s+(\d+)%', rem)
#         if m_above and p > float(m_above.group(1)):
#             return r["PO"]
#     return rules_list[0]["PO"]


# # ─────────────────────────────────────────────────────────────────────────────
# #  INSURER MATCHING HELPER
# # ─────────────────────────────────────────────────────────────────────────────

# def insurer_matches(rule_insurer_str, ins):
#     """
#     Check if the row's insurer matches a rule's INSURER field.

#     Handles entries like "Tata- Comp" where the rule token contains the
#     insurer name plus extra words/punctuation (e.g. product suffix).

#     Strategy per comma-separated token in the rule:
#       1. Exact match after normalisation.
#       2. First word of token equals insurer  →  "TATA- COMP" first word = "TATA".
#       3. Insurer is a substring of the token as a final fallback.
#     """
#     ri = str(rule_insurer_str).strip().upper()
#     if ri == "ALL COMPANIES":
#         return True

#     ins_norm = ins.strip().upper()
#     for token in [x.strip().upper() for x in ri.split(",")]:
#         if ins_norm == token:
#             return True
#         # Strip punctuation so "TATA- COMP" → ["TATA", "COMP"]
#         token_words = re.sub(r'[^A-Z0-9 ]', ' ', token).split()
#         if token_words and token_words[0] == ins_norm:
#             return True
#         if ins_norm in token:
#             return True

#     return False


# def filter_by_insurer(rules, ins):
#     """
#     Return best matching rules for given insurer:
#     1. Specific match (not 'All Companies', not 'Rest of Companies')
#     2. 'All Companies'
#     3. 'Rest of Companies'
#     """
#     specific = [r for r in rules
#                 if str(r.get("INSURER","")).strip().upper() not in ("ALL COMPANIES","REST OF COMPANIES")
#                 and insurer_matches(r.get("INSURER",""), ins)]
#     if specific:
#         return specific

#     all_co = [r for r in rules if str(r.get("INSURER","")).strip().upper() == "ALL COMPANIES"]
#     if all_co:
#         return all_co

#     rest = [r for r in rules if str(r.get("INSURER","")).strip().upper() == "REST OF COMPANIES"]
#     return rest


# # ─────────────────────────────────────────────────────────────────────────────
# #  DETERMINE JSON SEGMENT from row data
# # ─────────────────────────────────────────────────────────────────────────────

# def get_json_lob_and_segment(sub_product_name, segment, vehicle_type_id,
#                               from_wt, to_wt, company_code, is_od):
#     """
#     Returns (lob, json_segment) tuple for rule lookup.
#     lob is the JSON LOB string.
#     json_segment is the JSON SEGMENT string.
#     Returns (None, None) if not mappable.
#     """
#     sp = str(sub_product_name).strip()
#     seg = str(segment).strip().upper()
#     vt_id = int(vehicle_type_id) if pd.notna(vehicle_type_id) else 0
#     ins_upper = str(company_code).strip().upper()

#     # ── TWO WHEELER ──────────────────────────────────────────────────────────
#     if sp == "Two Wheeler":
#         lob = "TW"
#         if is_od or "COMP" in seg or "SAOD" in seg or "OD" in seg:
#             return lob, "TW SAOD + COMP"
#         else:  # TP Only
#             return lob, "TW TP"

#     # ── PRIVATE CAR ──────────────────────────────────────────────────────────
#     if sp == "Private Car":
#         lob = "PVT CAR"
#         if is_od or "COMP" in seg or "SAOD" in seg or "OD" in seg:
#             return lob, "PVT CAR COMP + SAOD"
#         else:
#             return lob, "PVT CAR TP"

#     # ── PASSENGER VEHICLE ────────────────────────────────────────────────────
#     if sp == "Passenger Vehicle":
#         # TAXI
#         if vt_id in TAXI_VEHICLE_IDS:
#             return "TAXI", "TAXI"

#         # STAFF BUS
#         if vt_id in STAFF_BUS_IDS:
#             return "BUS", "STAFF BUS"

#         # SCHOOL BUS
#         if vt_id in SCHOOL_BUS_IDS:
#             return "BUS", "SCHOOL BUS"

#         # ROUTE/PASSENGER BUS → treat as STAFF BUS (general BUS)
#         if vt_id in ROUTE_BUS_IDS:
#             return "BUS", "STAFF BUS"

#         # 3-wheelers (Auto, Electric Rikshaw, PCV 2W etc.)
#         if vt_id in PCV_3W_IDS:
#             return "GCV, PCV 3W", "All GVW & PCV 3W, GCV 3W"

#         # Tempo Traveller — treat as staff bus
#         return "BUS", "STAFF BUS"

#     # ── GOODS VEHICLE ────────────────────────────────────────────────────────
#     if sp == "Goods Vehicle":
#         from_w = float(from_wt) if pd.notna(from_wt) else 0
#         to_w   = float(to_wt)   if pd.notna(to_wt)   else 99999

#         # 3-Wheeler goods
#         if vt_id in GCV_3W_IDS:
#             return "GCV, PCV 3W", "All GVW & PCV 3W, GCV 3W"

#         # Upto 2.5T GVW + special insurers
#         if from_w == 0 and to_w <= 2500 and ins_upper in SPECIAL_GCV_INSURERS:
#             return "GCV, PCV 3W", "Upto 2.5 GVW"

#         # Everything else (inc. upto 2.5T with other insurers)
#         return "GCV, PCV 3W", "All GVW & PCV 3W, GCV 3W"

#     # ── MISCELLANEOUS VEHICLE (Tractor, Harvester etc.) ───────────────────────
#     if sp == "Miscellaneous Vehicle":
#         return "MISD", "Misd, Tractor"

#     return None, None


# # ─────────────────────────────────────────────────────────────────────────────
# #  CORE FORMULA ENGINE
# # ─────────────────────────────────────────────────────────────────────────────

# def compute_payout(payin, sub_product_name, segment, company_code,
#                    rto_group_name, vehicle_type_id, from_wt, to_wt,
#                    rules, is_od=True, rto_lookup=None):
#     """
#     Returns (payout_value, explanation_dict).
#     explanation_dict has keys: lob, segment, insurer_matched, po_formula,
#     remarks_slab, odisha_deduction, calculation_note
#     """
#     explanation = {
#         "lob": "",
#         "segment": "",
#         "insurer_matched": "",
#         "po_formula": "",
#         "remarks_slab": "",
#         "odisha_deduction": "",
#         "calculation_note": "",
#     }

#     try:
#         p = float(payin)
#     except (TypeError, ValueError):
#         explanation["calculation_note"] = "Invalid payin value"
#         return 0.0, explanation
#     if p == 0:
#         explanation["calculation_note"] = "Payin is 0 — no payout"
#         return 0.0, explanation

#     ins = str(company_code).strip().upper() if pd.notna(company_code) else ""
#     raw_loc = str(rto_group_name).strip() if pd.notna(rto_group_name) else ""
#     loc = resolve_state(raw_loc, rto_lookup or {})

#     lob, json_seg = get_json_lob_and_segment(
#         sub_product_name, segment, vehicle_type_id, from_wt, to_wt, company_code, is_od
#     )

#     explanation["lob"]     = lob     if lob     else "NOT MAPPED"
#     explanation["segment"] = json_seg if json_seg else "NOT MAPPED"

#     if lob is None:
#         p_out = floor_payout(p * 0.90)
#         explanation["po_formula"]        = "90% of Payin (fallback)"
#         explanation["insurer_matched"]   = "N/A"
#         explanation["remarks_slab"]      = "N/A"
#         explanation["calculation_note"]  = f"No LOB mapping found. Fallback: floor({p} × 0.90) = {p_out}"
#     else:
#         seg_rules      = [r for r in rules if r.get("LOB") == lob and r.get("SEGMENT") == json_seg]
#         candidate_rules = filter_by_insurer(seg_rules, ins)
#         selected_po    = select_po(candidate_rules, p)

#         if candidate_rules:
#             explanation["insurer_matched"] = str(candidate_rules[0].get("INSURER", ""))
#             # Find which slab was picked
#             for r in candidate_rules:
#                 rem = str(r.get("REMARKS","")).upper().strip()
#                 if not rem or rem == "NIL" or rem == "ALL FUEL":
#                     explanation["remarks_slab"] = r.get("REMARKS","NIL")
#                     break
#                 m_below = re.search(r'BELOW\s+(\d+)%', rem)
#                 if m_below and p <= float(m_below.group(1)):
#                     explanation["remarks_slab"] = r.get("REMARKS","")
#                     break
#                 m_range = re.search(r'(\d+)%\s+TO\s+(\d+)%', rem)
#                 if m_range:
#                     lo, hi = float(m_range.group(1)), float(m_range.group(2))
#                     if lo <= p <= hi:
#                         explanation["remarks_slab"] = r.get("REMARKS","")
#                         break
#                 m_above = re.search(r'ABOVE\s+(\d+)%', rem)
#                 if m_above and p > float(m_above.group(1)):
#                     explanation["remarks_slab"] = r.get("REMARKS","")
#                     break
#         else:
#             explanation["insurer_matched"] = "No matching rule"
#             explanation["remarks_slab"]    = "N/A"

#         if selected_po is None:
#             p_out = floor_payout(p * 0.90)
#             explanation["po_formula"]       = "90% of Payin (fallback — no rule matched)"
#             explanation["calculation_note"] = f"Fallback: floor({p} × 0.90) = {p_out}"
#         else:
#             explanation["po_formula"] = selected_po
#             p_out = parse_po_to_payout(selected_po, p)
#             # Build human-readable calculation note
#             po_up = str(selected_po).strip().upper()
#             if re.search(r'\d+%\s*OF PAYIN', po_up):
#                 pct = float(re.search(r'(\d+(?:\.\d+)?)%', po_up).group(1))
#                 explanation["calculation_note"] = f"floor({p} × {pct}/100) = {p_out}"
#             elif "PAYIN + 1" in po_up:
#                 explanation["calculation_note"] = f"floor({p} + 1) = {p_out}"
#             elif "LESS 2% OF PAYIN" in po_up:
#                 explanation["calculation_note"] = f"floor({p} - 2) = {p_out}"
#             elif re.fullmatch(r'-(\d+(?:\.\d+)?)%', po_up):
#                 ded = float(re.search(r'-(\d+(?:\.\d+)?)%', po_up).group(1))
#                 explanation["calculation_note"] = f"floor({p} - {ded}) = {p_out}"
#             elif re.search(r'(\d+(?:\.\d+)?)%\s*PO', po_up):
#                 explanation["calculation_note"] = f"Fixed PO = {p_out}"
#             else:
#                 explanation["calculation_note"] = f"= {p_out}"

#     # ── ODISHA OVERRIDE (additional deduction on top) ─────────────────────────
#     if loc == "ODISHA":
#         odisha_rules = [r for r in rules
#                         if r.get("LOCATION") == "ODISHA"
#                         and r.get("SEGMENT") == "ALL SEGMENT"
#                         and str(r.get("INSURER","")).strip().upper() == "ALL COMPANIES"]
#         ded_po = select_po(odisha_rules, p)
#         if ded_po:
#             ded_str = str(ded_po).strip().upper()
#             m = re.search(r'-(\d+(?:\.\d+)?)%', ded_str)
#             if m:
#                 ded = float(m.group(1))
#                 p_out = floor_payout(p - ded)   # slab on payin, deduction on payin
#                 explanation["odisha_deduction"] = (
#                     f"Odisha override ({ded_po}): floor(payin {p} - {ded}) = {p_out}"
#                 )

#     result = max(0.0, p_out)
#     return result, explanation


# # ─────────────────────────────────────────────────────────────────────────────
# #  MAIN
# # ─────────────────────────────────────────────────────────────────────────────

# def main():
#     print("\n" + "="*70)
#     print("  Payin-Config — Payout Recalculator")
#     print("="*70)

#     json_path   = input("\nEnter path to payout_rules.json : ").strip().strip('"')
#     input_path  = input("Enter path to PayinConfig.xlsx  : ").strip().strip('"')
#     master_path = input("Enter path to Master.xlsx        : ").strip().strip('"')
#     output_path = input("Enter output file path (blank=auto): ").strip().strip('"')

#     try:
#         with open(json_path) as f:
#             rules = json.load(f)
#         print(f"  Loaded {len(rules)} rules from {json_path}")
#     except Exception as e:
#         print(f"\n[ERROR] Failed to load JSON: {e}"); return

#     rto_lookup = build_rto_state_lookup(master_path)

#     if not output_path:
#         base, ext = os.path.splitext(input_path)
#         ts = datetime.now().strftime("%Y%m%d_%H%M%S")
#         output_path = f"{base}_recalculated_{ts}{ext}"

#     print(f"\n  Reading: {input_path}")
#     df = pd.read_excel(input_path)
#     df.columns = [c.strip() for c in df.columns]
#     total = len(df)
#     print(f"  Rows   : {total}")

#     required = ['payin_od_rate','payin_tp_rate','payout_od_rate','payout_tp_rate',
#                 'sub_product_name','segment','company_code','rto_group_name',
#                 'vehicle_type_id','from_weightage_kg','to_weightage_kg']
#     missing = [c for c in required if c not in df.columns]
#     if missing:
#         print(f"\n[ERROR] Missing columns: {missing}"); return

#     new_od, new_tp = [], []
#     changed_od = changed_tp = processed_od = processed_tp = 0

#     # Explanation column lists
#     od_lob, od_seg, od_ins, od_po, od_slab, od_odisha, od_note = [], [], [], [], [], [], []
#     tp_lob, tp_seg, tp_ins, tp_po, tp_slab, tp_odisha, tp_note = [], [], [], [], [], [], []

#     for _, row in df.iterrows():
#         sp   = row['sub_product_name']
#         seg  = row['segment']
#         ins  = row['company_code']
#         loc  = row['rto_group_name']
#         vt   = row['vehicle_type_id']
#         f_wt = row['from_weightage_kg']
#         t_wt = row['to_weightage_kg']

#         # OD
#         payin_od = row['payin_od_rate']
#         old_od   = row['payout_od_rate']
#         if pd.notna(payin_od) and float(payin_od) != 0:
#             processed_od += 1
#             calc_od, expl_od = compute_payout(payin_od, sp, seg, ins, loc, vt, f_wt, t_wt, rules, is_od=True, rto_lookup=rto_lookup)
#             new_od.append(calc_od)
#             if abs(float(old_od) - calc_od) > 0.001:
#                 changed_od += 1
#         else:
#             calc_od = 0.0 if pd.isna(old_od) else old_od
#             new_od.append(calc_od)
#             expl_od = {"lob":"","segment":"","insurer_matched":"","po_formula":"",
#                        "remarks_slab":"","odisha_deduction":"","calculation_note":"Payin is 0"}
#         od_lob.append(expl_od["lob"])
#         od_seg.append(expl_od["segment"])
#         od_ins.append(expl_od["insurer_matched"])
#         od_po.append(expl_od["po_formula"])
#         od_slab.append(expl_od["remarks_slab"])
#         od_odisha.append(expl_od["odisha_deduction"])
#         od_note.append(expl_od["calculation_note"])

#         # TP
#         payin_tp = row['payin_tp_rate']
#         old_tp   = row['payout_tp_rate']
#         if pd.notna(payin_tp) and float(payin_tp) != 0:
#             processed_tp += 1
#             calc_tp, expl_tp = compute_payout(payin_tp, sp, seg, ins, loc, vt, f_wt, t_wt, rules, is_od=False, rto_lookup=rto_lookup)
#             new_tp.append(calc_tp)
#             if abs(float(old_tp) - calc_tp) > 0.001:
#                 changed_tp += 1
#         else:
#             calc_tp = 0.0 if pd.isna(old_tp) else old_tp
#             new_tp.append(calc_tp)
#             expl_tp = {"lob":"","segment":"","insurer_matched":"","po_formula":"",
#                        "remarks_slab":"","odisha_deduction":"","calculation_note":"Payin is 0"}
#         tp_lob.append(expl_tp["lob"])
#         tp_seg.append(expl_tp["segment"])
#         tp_ins.append(expl_tp["insurer_matched"])
#         tp_po.append(expl_tp["po_formula"])
#         tp_slab.append(expl_tp["remarks_slab"])
#         tp_odisha.append(expl_tp["odisha_deduction"])
#         tp_note.append(expl_tp["calculation_note"])

#     df['payout_od_rate'] = new_od
#     df['payout_tp_rate'] = new_tp

#     # Append OD explanation columns
#     df['od_rule_lob']            = od_lob
#     df['od_rule_segment']        = od_seg
#     df['od_rule_insurer']        = od_ins
#     df['od_rule_po_formula']     = od_po
#     df['od_rule_slab']           = od_slab
#     df['od_rule_odisha']         = od_odisha
#     df['od_rule_calculation']    = od_note
    
#     # Append TP explanation columns 
#     df['tp_rule_lob']            = tp_lob
#     df['tp_rule_segment']        = tp_seg
#     df['tp_rule_insurer']        = tp_ins
#     df['tp_rule_po_formula']     = tp_po
#     df['tp_rule_slab']           = tp_slab
#     df['tp_rule_odisha']         = tp_odisha
#     df['tp_rule_calculation']    = tp_note

#     df.to_excel(output_path, index=False)

#     print(f"\n{'='*70}")
#     print(f"  COMPLETED")
#     print(f"  Total rows           : {total}")
#     print(f"  OD rows recalculated : {processed_od}   changed: {changed_od}")
#     print(f"  TP rows recalculated : {processed_tp}   changed: {changed_tp}")
#     print(f"  Output saved to      : {output_path}")
#     print(f"{'='*70}")

#     # Sample preview
#     sample_cols = ['sub_product_name','segment','company_code','rto_group_name',
#                    'vehicle_type_id','payin_od_rate','payout_od_rate',
#                    'payin_tp_rate','payout_tp_rate']
#     sample = df[df['payin_od_rate'] > 0].head(15)[sample_cols]
#     print("\n  Sample output (first 15 non-zero OD rows):\n")
#     print(sample.to_string(index=False))
#     print()


# if __name__ == "__main__":
#     main()

"""
Payin-Config — Payout Recalculator
===================================
Reads PayinConfig.xlsx + master Vehicle Type table (hardcoded from provided data)
and recomputes payout_od_rate / payout_tp_rate using JSON payout rules.

Usage:
    python recalculate_payout.py
"""

import pandas as pd
import math
import os
import json
import re
from datetime import datetime

# ─────────────────────────────────────────────────────────────────────────────
#  MASTER — Vehicle Type lookup (from master file Vehicle type worksheet)
# ─────────────────────────────────────────────────────────────────────────────

VEHICLE_TYPE_MASTER = {
    1:  {"vehicle_type": "Agriculture Tractor",          "sub_product_name": "Miscellaneous Vehicle"},
    2:  {"vehicle_type": "Non Tractor",                  "sub_product_name": "Miscellaneous Vehicle"},
    3:  {"vehicle_type": "Truck",                        "sub_product_name": "Goods Vehicle"},
    4:  {"vehicle_type": "Good Carring Tractor",         "sub_product_name": "Goods Vehicle"},
    5:  {"vehicle_type": "Tanker",                       "sub_product_name": "Goods Vehicle"},
    6:  {"vehicle_type": "Pickup",                       "sub_product_name": "Goods Vehicle"},
    7:  {"vehicle_type": "GCV 3W Delivery Van",          "sub_product_name": "Goods Vehicle"},
    8:  {"vehicle_type": "Taxi_CAB",                     "sub_product_name": "Passenger Vehicle"},
    9:  {"vehicle_type": "Electric Rikshaw",             "sub_product_name": "Passenger Vehicle"},
    10: {"vehicle_type": "Tempo Traveller",              "sub_product_name": "Passenger Vehicle"},
    11: {"vehicle_type": "School Bus",                   "sub_product_name": "Passenger Vehicle"},
    12: {"vehicle_type": "Passanger Bus",                "sub_product_name": "Passenger Vehicle"},
    13: {"vehicle_type": "Auto rikshaw",                 "sub_product_name": "Passenger Vehicle"},
    14: {"vehicle_type": "3W Tipper",                    "sub_product_name": "Goods Vehicle"},
    15: {"vehicle_type": "PCV 2W",                       "sub_product_name": "Passenger Vehicle"},
    16: {"vehicle_type": "GCV 2W",                       "sub_product_name": "Goods Vehicle"},
    17: {"vehicle_type": "TW Scooter",                   "sub_product_name": "Two Wheeler"},
    18: {"vehicle_type": "TW Bike",                      "sub_product_name": "Two Wheeler"},
    19: {"vehicle_type": "Private Car",                  "sub_product_name": "Private Car"},
    20: {"vehicle_type": "TW Electric Bike",             "sub_product_name": "Two Wheeler"},
    21: {"vehicle_type": "Electric GCV 3W Delivery Van", "sub_product_name": "Goods Vehicle"},
    22: {"vehicle_type": "Private Car Electric",         "sub_product_name": "Private Car"},
    23: {"vehicle_type": "Trailer",                      "sub_product_name": "Goods Vehicle"},
    24: {"vehicle_type": "Electric Pickup",              "sub_product_name": "Goods Vehicle"},
    25: {"vehicle_type": "Tipper",                       "sub_product_name": "Goods Vehicle"},
    26: {"vehicle_type": "TW Electric Scooter",          "sub_product_name": "Two Wheeler"},
    27: {"vehicle_type": "PC Petrol / Electric Hybrid",  "sub_product_name": "Private Car"},
    28: {"vehicle_type": "Staff Bus",                    "sub_product_name": "Passenger Vehicle"},
    29: {"vehicle_type": "Agriculture Harvester",        "sub_product_name": "Miscellaneous Vehicle"},
    30: {"vehicle_type": "Electric PCV 2W",              "sub_product_name": "Passenger Vehicle"},
    31: {"vehicle_type": "Electric Taxi_CAB",            "sub_product_name": "Passenger Vehicle"},
    32: {"vehicle_type": "Route Bus",                    "sub_product_name": "Passenger Vehicle"},
}

# IDs considered as TAXI
TAXI_VEHICLE_IDS = {8, 31}  # Taxi_CAB, Electric Taxi_CAB

# IDs considered as STAFF BUS
STAFF_BUS_IDS = {28}  # Staff Bus

# IDs considered as SCHOOL BUS
SCHOOL_BUS_IDS = {11}  # School Bus

# IDs considered as BUS (any bus — route/passenger)
ROUTE_BUS_IDS = {12, 32}  # Passanger Bus, Route Bus

# IDs considered as GCV 3-Wheeler goods
GCV_3W_IDS = {7, 14, 21}  # GCV 3W Delivery Van, 3W Tipper, Electric GCV 3W

# IDs considered as Passenger 3-Wheeler (auto etc.)
PCV_3W_IDS = {9, 13, 15, 30}  # Electric Rikshaw, Auto rikshaw, PCV 2W, Electric PCV 2W

# Insurers to match for Upto 2.5 GVW special rule
SPECIAL_GCV_INSURERS = {"RELIANCE", "SBI"}

# =============================================================================
#  RTO MAPPING LOOKUP — rto_group_id -> is_odisha flag
#  Uses PayInRTOMappingDump file.
#  Logic: if ANY rto_code for a given rto_group_id starts with "OD" or "OR"
#         -> that group is in Odisha.
# =============================================================================

def build_rto_odisha_lookup(rto_mapping_path):
    """
    Reads PayInRTOMappingDump.xlsx and builds a set of rto_group_ids
    that belong to Odisha (i.e. have at least one rto_code starting with OD or OR).
    Returns a set of integer rto_group_ids.
    """
    try:
        rto_df = pd.read_excel(rto_mapping_path)
        rto_df.columns = [c.strip() for c in rto_df.columns]
    except Exception as e:
        print(f"  [WARN] Could not load RTO Mapping file: {e}")
        return set()

    odisha_ids = set()
    for _, row in rto_df.iterrows():
        code = str(row.get('rto_code', '')).strip().upper()
        if code.startswith('OD') or code.startswith('OR'):
            try:
                odisha_ids.add(int(row['rto_group_id']))
            except (ValueError, TypeError):
                pass

    print(f"  RTO Mapping loaded: {len(odisha_ids)} Odisha rto_group_ids found.")
    return odisha_ids


def is_odisha(rto_group_id, odisha_ids):
    """Return True if the rto_group_id belongs to Odisha."""
    try:
        return int(rto_group_id) in odisha_ids
    except (ValueError, TypeError):
        return False


# ─


# ─────────────────────────────────────────────────────────────────────────────
#  FLOOR
# ─────────────────────────────────────────────────────────────────────────────

def floor_payout(value):
    return float(math.floor(float(value)))


# ─────────────────────────────────────────────────────────────────────────────
#  PO STRING PARSER
# ─────────────────────────────────────────────────────────────────────────────

def parse_po_to_payout(po_str, p):
    po_str = str(po_str).strip().upper()

    if re.search(r'\d+%\s*OF PAYIN', po_str):
        percent = float(re.search(r'(\d+(?:\.\d+)?)%', po_str).group(1))
        return floor_payout(p * (percent / 100))

    if "PAYIN + 1" in po_str:
        return floor_payout(p + 1)

    if "LESS 2% OF PAYIN" in po_str:
        return floor_payout(p - 2)

    # "-3%", "-4%", "-5%"
    m = re.fullmatch(r'-(\d+(?:\.\d+)?)%', po_str)
    if m:
        ded = float(m.group(1))
        return floor_payout(p - ded)

    # "21% PO" — fixed payout
    m = re.search(r'(\d+(?:\.\d+)?)%\s*PO', po_str)
    if m:
        return floor_payout(float(m.group(1)))

    # Fallback
    return floor_payout(p * 0.90)


# ─────────────────────────────────────────────────────────────────────────────
#  RULE SELECTION — slab-based matching on REMARKS
# ─────────────────────────────────────────────────────────────────────────────

def select_po(rules_list, p):
    if not rules_list:
        return None
    for r in rules_list:
        rem = str(r.get("REMARKS", "")).upper().strip()
        if not rem or rem == "NIL" or rem == "ALL FUEL":
            return r["PO"]
        m_below = re.search(r'BELOW\s+(\d+)%', rem)
        if m_below and p <= float(m_below.group(1)):
            return r["PO"]
        m_range = re.search(r'(\d+)%\s+TO\s+(\d+)%', rem)
        if m_range:
            lo, hi = float(m_range.group(1)), float(m_range.group(2))
            if lo <= p <= hi:
                return r["PO"]
        m_above = re.search(r'ABOVE\s+(\d+)%', rem)
        if m_above and p > float(m_above.group(1)):
            return r["PO"]
    return rules_list[0]["PO"]


# ─────────────────────────────────────────────────────────────────────────────
#  INSURER MATCHING HELPER
# ─────────────────────────────────────────────────────────────────────────────

def insurer_matches(rule_insurer_str, ins):
    """
    Check if the row's insurer matches a rule's INSURER field.

    Handles entries like "Tata- Comp" where the rule token contains the
    insurer name plus extra words/punctuation (e.g. product suffix).

    Strategy per comma-separated token in the rule:
      1. Exact match after normalisation.
      2. First word of token equals insurer  →  "TATA- COMP" first word = "TATA".
      3. Insurer is a substring of the token as a final fallback.
    """
    ri = str(rule_insurer_str).strip().upper()
    if ri == "ALL COMPANIES":
        return True

    ins_norm = ins.strip().upper()
    for token in [x.strip().upper() for x in ri.split(",")]:
        if ins_norm == token:
            return True
        # Strip punctuation so "TATA- COMP" → ["TATA", "COMP"]
        token_words = re.sub(r'[^A-Z0-9 ]', ' ', token).split()
        if token_words and token_words[0] == ins_norm:
            return True
        if ins_norm in token:
            return True

    return False


def filter_by_insurer(rules, ins):
    """
    Return best matching rules for given insurer:
    1. Specific match (not 'All Companies', not 'Rest of Companies')
    2. 'All Companies'
    3. 'Rest of Companies'
    """
    specific = [r for r in rules
                if str(r.get("INSURER","")).strip().upper() not in ("ALL COMPANIES","REST OF COMPANIES")
                and insurer_matches(r.get("INSURER",""), ins)]
    if specific:
        return specific

    all_co = [r for r in rules if str(r.get("INSURER","")).strip().upper() == "ALL COMPANIES"]
    if all_co:
        return all_co

    rest = [r for r in rules if str(r.get("INSURER","")).strip().upper() == "REST OF COMPANIES"]
    return rest


# ─────────────────────────────────────────────────────────────────────────────
#  DETERMINE JSON SEGMENT from row data
# ─────────────────────────────────────────────────────────────────────────────

def get_json_lob_and_segment(sub_product_name, segment, vehicle_type_id,
                              from_wt, to_wt, company_code, is_od):
    """
    Returns (lob, json_segment) tuple for rule lookup.
    lob is the JSON LOB string.
    json_segment is the JSON SEGMENT string.
    Returns (None, None) if not mappable.
    """
    sp = str(sub_product_name).strip()
    seg = str(segment).strip().upper()
    vt_id = int(vehicle_type_id) if pd.notna(vehicle_type_id) else 0
    ins_upper = str(company_code).strip().upper()

    # ── TWO WHEELER ──────────────────────────────────────────────────────────
    if sp == "Two Wheeler":
        lob = "TW"
        if is_od or "COMP" in seg or "SAOD" in seg or "OD" in seg:
            return lob, "TW SAOD + COMP"
        else:  # TP Only
            return lob, "TW TP"

    # ── PRIVATE CAR ──────────────────────────────────────────────────────────
    if sp == "Private Car":
        lob = "PVT CAR"
        if is_od or "COMP" in seg or "SAOD" in seg or "OD" in seg:
            return lob, "PVT CAR COMP + SAOD"
        else:
            return lob, "PVT CAR TP"

    # ── PASSENGER VEHICLE ────────────────────────────────────────────────────
    if sp == "Passenger Vehicle":
        # TAXI
        if vt_id in TAXI_VEHICLE_IDS:
            return "TAXI", "TAXI"

        # STAFF BUS
        if vt_id in STAFF_BUS_IDS:
            return "BUS", "STAFF BUS"

        # SCHOOL BUS
        if vt_id in SCHOOL_BUS_IDS:
            return "BUS", "SCHOOL BUS"

        # ROUTE/PASSENGER BUS → treat as STAFF BUS (general BUS)
        if vt_id in ROUTE_BUS_IDS:
            return "BUS", "STAFF BUS"

        # 3-wheelers (Auto, Electric Rikshaw, PCV 2W etc.)
        if vt_id in PCV_3W_IDS:
            return "GCV, PCV 3W", "All GVW & PCV 3W, GCV 3W"

        # Tempo Traveller — treat as staff bus
        return "BUS", "STAFF BUS"

    # ── GOODS VEHICLE ────────────────────────────────────────────────────────
    if sp == "Goods Vehicle":
        from_w = float(from_wt) if pd.notna(from_wt) else 0
        to_w   = float(to_wt)   if pd.notna(to_wt)   else 99999

        # 3-Wheeler goods
        if vt_id in GCV_3W_IDS:
            return "GCV, PCV 3W", "All GVW & PCV 3W, GCV 3W"

        # Upto 2.5T GVW + special insurers
        if from_w == 0 and to_w <= 2500 and ins_upper in SPECIAL_GCV_INSURERS:
            return "GCV, PCV 3W", "Upto 2.5 GVW"

        # Everything else (inc. upto 2.5T with other insurers)
        return "GCV, PCV 3W", "All GVW & PCV 3W, GCV 3W"

    # ── MISCELLANEOUS VEHICLE (Tractor, Harvester etc.) ───────────────────────
    if sp == "Miscellaneous Vehicle":
        return "MISD", "Misd, Tractor"

    return None, None


# ─────────────────────────────────────────────────────────────────────────────
#  CORE FORMULA ENGINE
# ─────────────────────────────────────────────────────────────────────────────

def compute_payout(payin, sub_product_name, segment, company_code,
                   rto_group_name, vehicle_type_id, from_wt, to_wt,
                   rules, is_od=True, odisha_ids=None):
    """
    Returns (payout_value, explanation_dict).
    explanation_dict has keys: lob, segment, insurer_matched, po_formula,
    remarks_slab, odisha_deduction, calculation_note
    """
    explanation = {
        "lob": "",
        "segment": "",
        "insurer_matched": "",
        "po_formula": "",
        "remarks_slab": "",
        "odisha_deduction": "",
        "calculation_note": "",
    }

    try:
        p = float(payin)
    except (TypeError, ValueError):
        explanation["calculation_note"] = "Invalid payin value"
        return 0.0, explanation
    if p == 0:
        explanation["calculation_note"] = "Payin is 0 — no payout"
        return 0.0, explanation

    ins = str(company_code).strip().upper() if pd.notna(company_code) else ""
    # Determine if this row is Odisha using rto_group_id (passed as rto_group_name param)
    loc_is_odisha = is_odisha(rto_group_name, odisha_ids or set())

    lob, json_seg = get_json_lob_and_segment(
        sub_product_name, segment, vehicle_type_id, from_wt, to_wt, company_code, is_od
    )

    explanation["lob"]     = lob     if lob     else "NOT MAPPED"
    explanation["segment"] = json_seg if json_seg else "NOT MAPPED"

    if lob is None:
        p_out = floor_payout(p * 0.90)
        explanation["po_formula"]        = "90% of Payin (fallback)"
        explanation["insurer_matched"]   = "N/A"
        explanation["remarks_slab"]      = "N/A"
        explanation["calculation_note"]  = f"No LOB mapping found. Fallback: floor({p} × 0.90) = {p_out}"
    else:
        seg_rules      = [r for r in rules if r.get("LOB") == lob and r.get("SEGMENT") == json_seg]
        candidate_rules = filter_by_insurer(seg_rules, ins)
        selected_po    = select_po(candidate_rules, p)

        if candidate_rules:
            explanation["insurer_matched"] = str(candidate_rules[0].get("INSURER", ""))
            # Find which slab was picked
            for r in candidate_rules:
                rem = str(r.get("REMARKS","")).upper().strip()
                if not rem or rem == "NIL" or rem == "ALL FUEL":
                    explanation["remarks_slab"] = r.get("REMARKS","NIL")
                    break
                m_below = re.search(r'BELOW\s+(\d+)%', rem)
                if m_below and p <= float(m_below.group(1)):
                    explanation["remarks_slab"] = r.get("REMARKS","")
                    break
                m_range = re.search(r'(\d+)%\s+TO\s+(\d+)%', rem)
                if m_range:
                    lo, hi = float(m_range.group(1)), float(m_range.group(2))
                    if lo <= p <= hi:
                        explanation["remarks_slab"] = r.get("REMARKS","")
                        break
                m_above = re.search(r'ABOVE\s+(\d+)%', rem)
                if m_above and p > float(m_above.group(1)):
                    explanation["remarks_slab"] = r.get("REMARKS","")
                    break
        else:
            explanation["insurer_matched"] = "No matching rule"
            explanation["remarks_slab"]    = "N/A"

        if selected_po is None:
            p_out = floor_payout(p * 0.90)
            explanation["po_formula"]       = "90% of Payin (fallback — no rule matched)"
            explanation["calculation_note"] = f"Fallback: floor({p} × 0.90) = {p_out}"
        else:
            explanation["po_formula"] = selected_po
            p_out = parse_po_to_payout(selected_po, p)
            # Build human-readable calculation note
            po_up = str(selected_po).strip().upper()
            if re.search(r'\d+%\s*OF PAYIN', po_up):
                pct = float(re.search(r'(\d+(?:\.\d+)?)%', po_up).group(1))
                explanation["calculation_note"] = f"floor({p} × {pct}/100) = {p_out}"
            elif "PAYIN + 1" in po_up:
                explanation["calculation_note"] = f"floor({p} + 1) = {p_out}"
            elif "LESS 2% OF PAYIN" in po_up:
                explanation["calculation_note"] = f"floor({p} - 2) = {p_out}"
            elif re.fullmatch(r'-(\d+(?:\.\d+)?)%', po_up):
                ded = float(re.search(r'-(\d+(?:\.\d+)?)%', po_up).group(1))
                explanation["calculation_note"] = f"floor({p} - {ded}) = {p_out}"
            elif re.search(r'(\d+(?:\.\d+)?)%\s*PO', po_up):
                explanation["calculation_note"] = f"Fixed PO = {p_out}"
            else:
                explanation["calculation_note"] = f"= {p_out}"

    # ── ODISHA OVERRIDE (additional deduction on top) ─────────────────────────
    if loc_is_odisha:
        odisha_rules = [r for r in rules
                        if r.get("LOCATION") == "ODISHA"
                        and r.get("SEGMENT") == "ALL SEGMENT"
                        and str(r.get("INSURER","")).strip().upper() == "ALL COMPANIES"]
        ded_po = select_po(odisha_rules, p)
        if ded_po:
            ded_str = str(ded_po).strip().upper()
            m = re.search(r'-(\d+(?:\.\d+)?)%', ded_str)
            if m:
                ded = float(m.group(1))
                p_out = floor_payout(p - ded)   # slab on payin, deduction on payin
                explanation["odisha_deduction"] = (
                    f"Odisha override ({ded_po}): floor(payin {p} - {ded}) = {p_out}"
                )

    result = max(0.0, p_out)
    return result, explanation


# ─────────────────────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    print("\n" + "="*70)
    print("  Payin-Config — Payout Recalculator")
    print("="*70)

    json_path   = input("\nEnter path to payout_rules.json : ").strip().strip('"')
    input_path  = input("Enter path to PayinConfig.xlsx  : ").strip().strip('"')
    rto_mapping_path = input("Enter path to PayInRTOMappingDump.xlsx: ").strip().strip('"')
    output_path = input("Enter output file path (blank=auto): ").strip().strip('"')

    try:
        with open(json_path) as f:
            rules = json.load(f)
        print(f"  Loaded {len(rules)} rules from {json_path}")
    except Exception as e:
        print(f"\n[ERROR] Failed to load JSON: {e}"); return

    odisha_ids = build_rto_odisha_lookup(rto_mapping_path)

    if not output_path:
        base, ext = os.path.splitext(input_path)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = f"{base}_recalculated_{ts}{ext}"

    print(f"\n  Reading: {input_path}")
    df = pd.read_excel(input_path)
    df.columns = [c.strip() for c in df.columns]
    total = len(df)
    print(f"  Rows   : {total}")

    required = ['payin_od_rate','payin_tp_rate','payout_od_rate','payout_tp_rate',
                'sub_product_name','segment','company_code','rto_group_id',
                'vehicle_type_id','from_weightage_kg','to_weightage_kg']
    missing = [c for c in required if c not in df.columns]
    if missing:
        print(f"\n[ERROR] Missing columns: {missing}"); return

    new_od, new_tp = [], []
    changed_od = changed_tp = processed_od = processed_tp = 0

    # Explanation column lists
    od_lob, od_seg, od_ins, od_po, od_slab, od_odisha, od_note = [], [], [], [], [], [], []
    tp_lob, tp_seg, tp_ins, tp_po, tp_slab, tp_odisha, tp_note = [], [], [], [], [], [], []

    for _, row in df.iterrows():
        sp   = row['sub_product_name']
        seg  = row['segment']
        ins  = row['company_code']
        vt   = row['vehicle_type_id']
        f_wt = row['from_weightage_kg']
        t_wt = row['to_weightage_kg']
        rto_grp_id = row['rto_group_id']

        # OD
        payin_od = row['payin_od_rate']
        old_od   = row['payout_od_rate']
        if pd.notna(payin_od) and float(payin_od) != 0:
            processed_od += 1
            calc_od, expl_od = compute_payout(payin_od, sp, seg, ins, rto_grp_id, vt, f_wt, t_wt, rules, is_od=True, odisha_ids=odisha_ids)
            new_od.append(calc_od)
            if abs(float(old_od) - calc_od) > 0.001:
                changed_od += 1
        else:
            calc_od = 0.0 if pd.isna(old_od) else old_od
            new_od.append(calc_od)
            expl_od = {"lob":"","segment":"","insurer_matched":"","po_formula":"",
                       "remarks_slab":"","odisha_deduction":"","calculation_note":"Payin is 0"}
        od_lob.append(expl_od["lob"])
        od_seg.append(expl_od["segment"])
        od_ins.append(expl_od["insurer_matched"])
        od_po.append(expl_od["po_formula"])
        od_slab.append(expl_od["remarks_slab"])
        od_odisha.append(expl_od["odisha_deduction"])
        od_note.append(expl_od["calculation_note"])

        # TP
        payin_tp = row['payin_tp_rate']
        old_tp   = row['payout_tp_rate']
        if pd.notna(payin_tp) and float(payin_tp) != 0:
            processed_tp += 1
            calc_tp, expl_tp = compute_payout(payin_tp, sp, seg, ins, rto_grp_id, vt, f_wt, t_wt, rules, is_od=False, odisha_ids=odisha_ids)
            new_tp.append(calc_tp)
            if abs(float(old_tp) - calc_tp) > 0.001:
                changed_tp += 1
        else:
            calc_tp = 0.0 if pd.isna(old_tp) else old_tp
            new_tp.append(calc_tp)
            expl_tp = {"lob":"","segment":"","insurer_matched":"","po_formula":"",
                       "remarks_slab":"","odisha_deduction":"","calculation_note":"Payin is 0"}
        tp_lob.append(expl_tp["lob"])
        tp_seg.append(expl_tp["segment"])
        tp_ins.append(expl_tp["insurer_matched"])
        tp_po.append(expl_tp["po_formula"])
        tp_slab.append(expl_tp["remarks_slab"])
        tp_odisha.append(expl_tp["odisha_deduction"])
        tp_note.append(expl_tp["calculation_note"])

    df['payout_od_rate'] = new_od
    df['payout_tp_rate'] = new_tp

    # Append OD explanation columns
    df['od_rule_lob']            = od_lob
    df['od_rule_segment']        = od_seg
    df['od_rule_insurer']        = od_ins
    df['od_rule_po_formula']     = od_po
    df['od_rule_slab']           = od_slab
    df['od_rule_odisha']         = od_odisha
    df['od_rule_calculation']    = od_note

    # Append TP explanation columns
    df['tp_rule_lob']            = tp_lob
    df['tp_rule_segment']        = tp_seg
    df['tp_rule_insurer']        = tp_ins
    df['tp_rule_po_formula']     = tp_po
    df['tp_rule_slab']           = tp_slab
    df['tp_rule_odisha']         = tp_odisha
    df['tp_rule_calculation']    = tp_note

    df.to_excel(output_path, index=False)

    print(f"\n{'='*70}")
    print(f"  COMPLETED")
    print(f"  Total rows           : {total}")
    print(f"  OD rows recalculated : {processed_od}   changed: {changed_od}")
    print(f"  TP rows recalculated : {processed_tp}   changed: {changed_tp}")
    print(f"  Output saved to      : {output_path}")
    print(f"{'='*70}")

    # Sample preview
    sample_cols = ['sub_product_name','segment','company_code','rto_group_id',
                   'rto_group_id','vehicle_type_id','payin_od_rate','payout_od_rate',
                   'payin_tp_rate','payout_tp_rate']
    sample = df[df['payin_od_rate'] > 0].head(15)[sample_cols]
    print("\n  Sample output (first 15 non-zero OD rows):\n")
    print(sample.to_string(index=False))
    print()


if __name__ == "__main__":
    main()
