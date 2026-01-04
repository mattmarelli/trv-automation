import pandas as pd
import xlsxwriter

from constants import fault_locations_table, fault_types_table, test_duty_table

def split_brk_data_by_first_to_clear(brk_data, trv_data):
    trv_flag = trv_data[["Run #", "Loc1/Rem2 First"]].copy()
    brk_tagged = brk_data.merge(trv_flag, on="Run #", how="inner")

    brk_local = brk_tagged[brk_tagged["Loc1/Rem2 First"] == 1].copy()
    brk_remote = brk_tagged[brk_tagged["Loc1/Rem2 First"] == 2].copy()

    return brk_local, brk_remote

def create_test_duty_phase_buckets(brk_data, test_duty_bucket_values):
    brk_data_by_test_duty_brackets = {
        "10%": {},
        "30%": {},
        "60%": {},
        "100%": {},
    }

    brk_data_by_test_duty_brackets["10%"]["a"] = brk_data[
        brk_data["BRK1A_RMS"].between(
            test_duty_bucket_values["10%"]["low"],
            test_duty_bucket_values["10%"]["high"],
            inclusive="left"
        )
    ]

    brk_data_by_test_duty_brackets["30%"]["a"] = brk_data[
        brk_data["BRK1A_RMS"].between(
            test_duty_bucket_values["30%"]["low"],
            test_duty_bucket_values["30%"]["high"],
            inclusive="left"
        )
    ]

    brk_data_by_test_duty_brackets["60%"]["a"] = brk_data[
        brk_data["BRK1A_RMS"].between(
            test_duty_bucket_values["60%"]["low"],
            test_duty_bucket_values["60%"]["high"],
            inclusive="left"
        )
    ]

    brk_data_by_test_duty_brackets["100%"]["a"] = brk_data[
        brk_data["BRK1A_RMS"].between(
            test_duty_bucket_values["100%"]["low"],
            test_duty_bucket_values["100%"]["high"],
            inclusive="left"
        )
    ]

    brk_data_by_test_duty_brackets["10%"]["b"] = brk_data[
        brk_data["BRK1B_RMS"].between(
            test_duty_bucket_values["10%"]["low"],
            test_duty_bucket_values["10%"]["high"],
            inclusive="left"
        )
    ]

    brk_data_by_test_duty_brackets["30%"]["b"] = brk_data[
        brk_data["BRK1B_RMS"].between(
            test_duty_bucket_values["30%"]["low"],
            test_duty_bucket_values["30%"]["high"],
            inclusive="left"
        )
    ]

    brk_data_by_test_duty_brackets["60%"]["b"] = brk_data[
        brk_data["BRK1B_RMS"].between(
            test_duty_bucket_values["60%"]["low"],
            test_duty_bucket_values["60%"]["high"],
            inclusive="left"
        )
    ]

    brk_data_by_test_duty_brackets["100%"]["b"] = brk_data[
        brk_data["BRK1B_RMS"].between(
            test_duty_bucket_values["100%"]["low"],
            test_duty_bucket_values["100%"]["high"],
            inclusive="left"
        )
    ]

    brk_data_by_test_duty_brackets["10%"]["c"] = brk_data[
        brk_data["BRK1C_RMS"].between(
            test_duty_bucket_values["10%"]["low"],
            test_duty_bucket_values["10%"]["high"],
            inclusive="left"
        )
    ]

    brk_data_by_test_duty_brackets["30%"]["c"] = brk_data[
        brk_data["BRK1C_RMS"].between(
            test_duty_bucket_values["30%"]["low"],
            test_duty_bucket_values["30%"]["high"],
            inclusive="left"
        )
    ]

    brk_data_by_test_duty_brackets["60%"]["c"] = brk_data[
        brk_data["BRK1C_RMS"].between(
            test_duty_bucket_values["60%"]["low"],
            test_duty_bucket_values["60%"]["high"],
            inclusive="left"
        )
    ]

    brk_data_by_test_duty_brackets["100%"]["c"] = brk_data[
        brk_data["BRK1C_RMS"].between(
            test_duty_bucket_values["100%"]["low"],
            test_duty_bucket_values["100%"]["high"],
            inclusive="left"
        )
    ]

    return brk_data_by_test_duty_brackets

def find_peaks(brk_bukets, trv_data):
    trv_data_10_a = trv_data.merge(brk_bukets["10%"]["a"][["Run #", "BRK1A_RMS"]], on="Run #", how="inner")
    a_10_kv_peak = trv_data_10_a.loc[trv_data_10_a["CB1_A_Peak(kV)"].idxmax()] if not trv_data_10_a.empty else pd.DataFrame()
    a_10_rrrv_peak = trv_data_10_a.loc[trv_data_10_a["CB1_A_RRRV(kV/u)"].idxmax()] if not trv_data_10_a.empty else pd.DataFrame()

    trv_data_10_b = trv_data.merge(brk_bukets["10%"]["b"][["Run #", "BRK1B_RMS"]], on="Run #", how="inner")
    b_10_kv_peak = trv_data_10_b.loc[trv_data_10_b["CB1_B_Peak(kV)"].idxmax()] if not trv_data_10_b.empty else pd.DataFrame()
    b_10_rrrv_peak = trv_data_10_b.loc[trv_data_10_b["CB1_B_RRRV(kV/u)"].idxmax()] if not trv_data_10_b.empty else pd.DataFrame()

    trv_data_10_c = trv_data.merge(brk_bukets["10%"]["c"][["Run #", "BRK1C_RMS"]], on="Run #", how="inner")
    c_10_kv_peak = trv_data_10_c.loc[trv_data_10_c["CB1_C_Peak(kV)"].idxmax()] if not trv_data_10_c.empty else pd.DataFrame()
    c_10_rrrv_peak = trv_data_10_c.loc[trv_data_10_c["CB1_C_RRRV(kV/u)"].idxmax()] if not trv_data_10_c.empty else pd.DataFrame()


    trv_data_30_a = trv_data.merge(brk_bukets["30%"]["a"][["Run #", "BRK1A_RMS"]], on="Run #", how="inner")
    a_30_kv_peak = trv_data_30_a.loc[trv_data_30_a["CB1_A_Peak(kV)"].idxmax()] if not trv_data_30_a.empty else pd.DataFrame()
    a_30_rrrv_peak = trv_data_30_a.loc[trv_data_30_a["CB1_A_RRRV(kV/u)"].idxmax()] if not trv_data_30_a.empty else pd.DataFrame()

    trv_data_30_b = trv_data.merge(brk_bukets["30%"]["b"][["Run #", "BRK1B_RMS"]], on="Run #", how="inner")
    b_30_kv_peak = trv_data_30_b.loc[trv_data_30_b["CB1_B_Peak(kV)"].idxmax()] if not trv_data_30_b.empty else pd.DataFrame()
    b_30_rrrv_peak = trv_data_30_b.loc[trv_data_30_b["CB1_B_RRRV(kV/u)"].idxmax()] if not trv_data_30_b.empty else pd.DataFrame()

    trv_data_30_c = trv_data.merge(brk_bukets["30%"]["c"][["Run #", "BRK1C_RMS"]], on="Run #", how="inner")
    c_30_kv_peak = trv_data_30_c.loc[trv_data_30_c["CB1_C_Peak(kV)"].idxmax()] if not trv_data_30_c.empty else pd.DataFrame()
    c_30_rrrv_peak = trv_data_30_c.loc[trv_data_30_c["CB1_C_RRRV(kV/u)"].idxmax()] if not trv_data_30_c.empty else pd.DataFrame()

    trv_data_60_a = trv_data.merge(brk_bukets["60%"]["a"][["Run #", "BRK1A_RMS"]], on="Run #", how="inner")
    a_60_kv_peak = trv_data_60_a.loc[trv_data_60_a["CB1_A_Peak(kV)"].idxmax()] if not trv_data_60_a.empty else pd.DataFrame()
    a_60_rrrv_peak = trv_data_60_a.loc[trv_data_60_a["CB1_A_RRRV(kV/u)"].idxmax()] if not trv_data_60_a.empty else pd.DataFrame()

    trv_data_60_b = trv_data.merge(brk_bukets["60%"]["b"][["Run #", "BRK1B_RMS"]], on="Run #", how="inner")
    b_60_kv_peak = trv_data_60_b.loc[trv_data_60_b["CB1_B_Peak(kV)"].idxmax()] if not trv_data_60_b.empty else pd.DataFrame()
    b_60_rrrv_peak = trv_data_60_b.loc[trv_data_60_b["CB1_B_RRRV(kV/u)"].idxmax()] if not trv_data_60_b.empty else pd.DataFrame()

    trv_data_60_c = trv_data.merge(brk_bukets["60%"]["c"][["Run #", "BRK1C_RMS"]], on="Run #", how="inner")
    c_60_kv_peak = trv_data_60_c.loc[trv_data_60_c["CB1_C_Peak(kV)"].idxmax()] if not trv_data_60_c.empty else pd.DataFrame()
    c_60_rrrv_peak = trv_data_60_c.loc[trv_data_60_c["CB1_C_RRRV(kV/u)"].idxmax()] if not trv_data_60_c.empty else pd.DataFrame()

    trv_data_100_a = trv_data.merge(brk_bukets["100%"]["a"][["Run #", "BRK1A_RMS"]], on="Run #", how="inner")
    a_100_kv_peak = trv_data_100_a.loc[trv_data_100_a["CB1_A_Peak(kV)"].idxmax()] if not trv_data_100_a.empty else pd.DataFrame()
    a_100_rrrv_peak = trv_data_100_a.loc[trv_data_100_a["CB1_A_RRRV(kV/u)"].idxmax()] if not trv_data_100_a.empty else pd.DataFrame()

    trv_data_100_b = trv_data.merge(brk_bukets["100%"]["b"][["Run #", "BRK1B_RMS"]], on="Run #", how="inner")
    b_100_kv_peak = trv_data_100_b.loc[trv_data_100_b["CB1_B_Peak(kV)"].idxmax()] if not trv_data_100_b.empty else pd.DataFrame()
    b_100_rrrv_peak = trv_data_100_b.loc[trv_data_100_b["CB1_B_RRRV(kV/u)"].idxmax()] if not trv_data_100_b.empty else pd.DataFrame()

    trv_data_100_c = trv_data.merge(brk_bukets["100%"]["c"][["Run #", "BRK1C_RMS"]], on="Run #", how="inner")
    c_100_kv_peak = trv_data_100_c.loc[trv_data_100_c["CB1_C_Peak(kV)"].idxmax()] if not trv_data_100_c.empty else pd.DataFrame()
    c_100_rrrv_peak = trv_data_100_c.loc[trv_data_100_c["CB1_C_RRRV(kV/u)"].idxmax()] if not trv_data_100_c.empty else pd.DataFrame()

    peaks = {
        "10%": {
            "a": {
                "kv": a_10_kv_peak,
                "rrrv": a_10_rrrv_peak,
            },
            "b": {
                "kv": b_10_kv_peak,
                "rrrv": b_10_rrrv_peak,
            },
            "c": {
                "kv": c_10_kv_peak,
                "rrrv": c_10_rrrv_peak,
            },
        },
        "30%": {
            "a": {
                "kv": a_30_kv_peak,
                "rrrv": a_30_rrrv_peak,
            },
            "b": {
                "kv": b_30_kv_peak,
                "rrrv": b_30_rrrv_peak,
            },
            "c": {
                "kv": c_30_kv_peak,
                "rrrv": c_30_rrrv_peak,
            },
        },
        "60%": {
            "a": {
                "kv": a_60_kv_peak,
                "rrrv": a_60_rrrv_peak,
            },
            "b": {
                "kv": b_60_kv_peak,
                "rrrv": b_60_rrrv_peak,
            },
            "c": {
                "kv": c_60_kv_peak,
                "rrrv": c_60_rrrv_peak,
            },
        },
        "100%": {
            "a": {
                "kv": a_100_kv_peak,
                "rrrv": a_100_rrrv_peak,
            },
            "b": {
                "kv": b_100_kv_peak,
                "rrrv": b_100_rrrv_peak,
            },
            "c": {
                "kv": c_100_kv_peak,
                "rrrv": c_100_rrrv_peak,
            },
        },
    }

    return peaks

def turn_peaks_into_output_rows(peaks):
    output_rows = {}
    maximum_output_rows = {}

    for peak_key, peak_values in peaks.items():
        output_rows[peak_key] = {
            "kv": {},
            "rrrv": {},
        }
        maximum_output_rows[peak_key] = {
            "kv": {},
            "rrrv": {},
        }
        maximum_kv_value = -1
        maximum_kv_output = None
        maximum_rrrv_value = -1
        maximum_rrrv_output = None
        for phase, phase_values in peak_values.items():
            if phase == "a":
                kv_string = "CB1_A_Peak(kV)"
                rrrv_string = "CB1_A_RRRV(kV/u)"
                rms_string = "BRK1A_RMS"
            elif phase == "b":
                kv_string = "CB1_B_Peak(kV)"
                rrrv_string = "CB1_B_RRRV(kV/u)"
                rms_string = "BRK1B_RMS"
            elif phase == "c":
                kv_string = "CB1_C_Peak(kV)"
                rrrv_string = "CB1_C_RRRV(kV/u)"
                rms_string = "BRK1C_RMS"

            if not phase_values["kv"].empty:
                fault_type = int(phase_values["kv"]["Fault_Type"])
                fault_type_string = fault_types_table[fault_type]
                fault_location = int(phase_values["kv"]["Fault_Location"])
                fault_location_string = fault_locations_table[fault_location]
                peak_kv = phase_values["kv"][kv_string]
                interrupting_current = phase_values["kv"][rms_string]
                output_value = {
                    "fault_type": fault_type_string,
                    "fault_location": fault_location_string,
                    "interrupting_current": interrupting_current,
                    "peak_value": peak_kv,
                }
                if peak_kv > maximum_kv_value:
                    maximum_kv_value = peak_kv
                    maximum_kv_output = output_value
            else:
                output_value = {
                    "fault_type": "N/A",
                    "fault_location": "N/A",
                    "interrupting_current": "N/A",
                    "peak_value": "N/A",
                }
            
            output_rows[peak_key]["kv"][phase] = output_value

            if not phase_values["rrrv"].empty:
                fault_type = int(phase_values["rrrv"]["Fault_Type"])
                fault_type_string = fault_types_table[fault_type]
                fault_location = int(phase_values["rrrv"]["Fault_Location"])
                fault_location_string = fault_locations_table[fault_location]
                peak_rrrv = phase_values["rrrv"][rrrv_string]
                interrupting_current = phase_values["rrrv"][rms_string]
                output_value = {
                    "fault_type": fault_type_string,
                    "fault_location": fault_location_string,
                    "interrupting_current": interrupting_current,
                    "peak_value": peak_rrrv,
                }
                if peak_rrrv > maximum_rrrv_value:
                    maximum_rrrv_value = peak_rrrv
                    maximum_rrrv_output = output_value
            else:
                output_value = {
                    "fault_type": "N/A",
                    "fault_location": "N/A",
                    "interrupting_current": "N/A",
                    "peak_value": "N/A",
                }
            
            output_rows[peak_key]["rrrv"][phase] = output_value
        
        if maximum_kv_output:
            maximum_output_rows[peak_key]["kv"] = maximum_kv_output
        else:
            maximum_output_rows[peak_key]["kv"] = {
                "fault_type": "N/A",
                "fault_location": "N/A",
                "interrupting_current": "N/A",
                "peak_value": "N/A",
            }

        if maximum_rrrv_output:
            maximum_output_rows[peak_key]["rrrv"] = maximum_rrrv_output
        else:
            maximum_output_rows[peak_key]["rrrv"] = {
                "fault_type": "N/A",
                "fault_location": "N/A",
                "interrupting_current": "N/A",
                "peak_value": "N/A",
            }

    return output_rows, maximum_output_rows


def create_output_file(local_peaks, remote_peaks, selected_breaker_voltage_class, local_station, remote_station, breaker_names):
    IEEE_standards = test_duty_table[selected_breaker_voltage_class]
    local_peaks_output, maximum_local_peaks_output = turn_peaks_into_output_rows(local_peaks)
    remote_peaks_output, maximum_remote_peaks_output = turn_peaks_into_output_rows(remote_peaks)

    breaker_names_for_file_name = breaker_names.replace("\\", "_").replace("/", "_")
    workbook = xlsxwriter.Workbook(f"{breaker_names_for_file_name} trv-automation results.xlsx")
    center_format = workbook.add_format({
        "align": "center",
        "valign": "vcenter",
        "border": 2,
        "text_wrap": True
    })
    center_format_bold = workbook.add_format({
        "align": "center",
        "valign": "vcenter",
        "border": 2,
        "bold": True,
        "text_wrap": True
    })
    center_format_red_text = workbook.add_format({
        "align": "center",
        "valign": "vcenter",
        "border": 2,
        "text_wrap": True,
        "font_color": "red"
    })
    center_format_gray_background = workbook.add_format({
        "align": "center",
        "valign": "vcenter",
        "bg_color": "#A3A3A3",
        "pattern": 1,
        "bold": True,
        "text_wrap": True,
        "border": 2
    })
    worksheet = workbook.add_worksheet()
    worksheet.set_column(first_col=0, last_col=0, width=14)
    worksheet.set_column(first_col=1, last_col=1, width=14)
    worksheet.set_column(first_col=5, last_col=5, width=16)
    worksheet.set_column(first_col=6, last_col=6, width=12)
    worksheet.set_column(first_col=8, last_col=8, width=14)
    worksheet.set_column(first_col=11, last_col=11, width=12)
    worksheet.set_column(first_col=13, last_col=13, width=14)
    
    table_header = f"{local_station} to {remote_station} {breaker_names} Breaker Set Results For Line Faults"
    worksheet.merge_range("A1:N1", table_header, center_format)

    worksheet.merge_range("A2:A6", "Breaker", center_format_gray_background)
    worksheet.merge_range("B2:B6", "First Breaker to Clear", center_format_gray_background)
    worksheet.merge_range("C2:C6", "Test Duty", center_format_gray_background)
    worksheet.merge_range("D2:D6", "Phase", center_format_gray_background)

    # TRV Controlling
    worksheet.merge_range("E2:I2", "TRV Controlling", center_format_gray_background)
    worksheet.merge_range("E3:F3", "Case", center_format_gray_background)
    worksheet.merge_range("E4:E6", "Fault Type", center_format_gray_background)
    worksheet.merge_range("F4:F6", "Fault Location*", center_format_gray_background)
    worksheet.merge_range("G3:G6", "Interrupting Current (kA)", center_format_gray_background)
    worksheet.merge_range("H3:H6", "Peak TRV** (kV)", center_format_gray_background)
    worksheet.merge_range("I3:I6", "IEEE Recommended Rating (kV)", center_format_gray_background)

    # RRRV Controlling
    worksheet.merge_range("J2:N2", "RRRV Controlling", center_format_gray_background)
    worksheet.merge_range("J3:K3", "Case", center_format_gray_background)
    worksheet.merge_range("J4:J6", "Fault Type", center_format_gray_background)
    worksheet.merge_range("K4:K6", "Fault Location*", center_format_gray_background)
    worksheet.merge_range("L3:L6", "Interrupting Current (kA)", center_format_gray_background)
    worksheet.merge_range("M3:M6", "Peak RRRV** (kV/μs)", center_format_gray_background)
    worksheet.merge_range("N3:N6", "IEEE Recommended Rating (kV/μs)", center_format_gray_background)
    
    
    worksheet.merge_range("A7:A30", breaker_names, center_format)
    worksheet.merge_range("B7:B18", f"{local_station} (Local)", center_format)
    worksheet.merge_range("B19:B30", f"{remote_station} (Remote)", center_format)

    # Test Duty percentages
    worksheet.merge_range("C7:C9", "10%", center_format_bold)
    worksheet.merge_range("C10:C12", "30%", center_format_bold)
    worksheet.merge_range("C13:C15", "60%", center_format_bold)
    worksheet.merge_range("C16:C18", "100%", center_format_bold)
    worksheet.merge_range("C19:C21", "10%", center_format_bold)
    worksheet.merge_range("C22:C24", "30%", center_format_bold)
    worksheet.merge_range("C25:C27", "60%", center_format_bold)
    worksheet.merge_range("C28:C30", "100%", center_format_bold)

    # Phase
    worksheet.write("D7", "A", center_format_gray_background)
    worksheet.write("D8", "B", center_format_gray_background)
    worksheet.write("D9", "C", center_format_gray_background)
    worksheet.write("D10", "A", center_format_gray_background)
    worksheet.write("D11", "B", center_format_gray_background)
    worksheet.write("D12", "C", center_format_gray_background)
    worksheet.write("D13", "A", center_format_gray_background)
    worksheet.write("D14", "B", center_format_gray_background)
    worksheet.write("D15", "C", center_format_gray_background)
    worksheet.write("D16", "A", center_format_gray_background)
    worksheet.write("D17", "B", center_format_gray_background)
    worksheet.write("D18", "C", center_format_gray_background)
    worksheet.write("D19", "A", center_format_gray_background)
    worksheet.write("D20", "B", center_format_gray_background)
    worksheet.write("D21", "C", center_format_gray_background)
    worksheet.write("D22", "A", center_format_gray_background)
    worksheet.write("D23", "B", center_format_gray_background)
    worksheet.write("D24", "C", center_format_gray_background)
    worksheet.write("D25", "A", center_format_gray_background)
    worksheet.write("D26", "B", center_format_gray_background)
    worksheet.write("D27", "C", center_format_gray_background)
    worksheet.write("D28", "A", center_format_gray_background)
    worksheet.write("D29", "B", center_format_gray_background)
    worksheet.write("D30", "C", center_format_gray_background)

    # Filling in the datatable
    write_output_values(worksheet, center_format, local_peaks_output, "10%", 7, 8, 9, IEEE_standards)
    write_output_values(worksheet, center_format, local_peaks_output, "30%", 10, 11, 12, IEEE_standards)
    write_output_values(worksheet, center_format, local_peaks_output, "60%", 13, 14, 15, IEEE_standards)
    write_output_values(worksheet, center_format, local_peaks_output, "100%", 16, 17, 18, IEEE_standards)
    write_output_values(worksheet, center_format, remote_peaks_output, "10%", 19, 20, 21, IEEE_standards)
    write_output_values(worksheet, center_format, remote_peaks_output, "30%", 22, 23, 24, IEEE_standards)
    write_output_values(worksheet, center_format, remote_peaks_output, "60%", 25, 26, 27, IEEE_standards)
    write_output_values(worksheet, center_format, remote_peaks_output, "100%", 28, 29, 30, IEEE_standards)


    # Smaller table at the bottom
    worksheet.merge_range("B32:N32", table_header, center_format)

    worksheet.merge_range("B33:B37", "Breaker", center_format_gray_background)
    worksheet.merge_range("C33:C37", "First Breaker to Clear", center_format_gray_background)
    worksheet.merge_range("D33:D37", "Test Duty", center_format_gray_background)

    # TRV Controlling
    worksheet.merge_range("E33:I33", "TRV Controlling", center_format_gray_background)
    worksheet.merge_range("E34:F34", "Case", center_format_gray_background)
    worksheet.merge_range("E35:E37", "Fault Type", center_format_gray_background)
    worksheet.merge_range("F35:F37", "Fault Location*", center_format_gray_background)
    worksheet.merge_range("G34:G37", "Interrupting Current (kA)", center_format_gray_background)
    worksheet.merge_range("H34:H37", "Peak TRV** (kV)", center_format_gray_background)
    worksheet.merge_range("I34:I37", "IEEE Recommended Rating (kV)", center_format_gray_background)

    # RRRV Controlling
    worksheet.merge_range("J33:N33", "RRRV Controlling", center_format_gray_background)
    worksheet.merge_range("J34:K34", "Case", center_format_gray_background)
    worksheet.merge_range("J35:J37", "Fault Type", center_format_gray_background)
    worksheet.merge_range("K35:K37", "Fault Location*", center_format_gray_background)
    worksheet.merge_range("L34:L37", "Interrupting Current (kA)", center_format_gray_background)
    worksheet.merge_range("M34:M37", "Peak RRRV** (kV/μs)", center_format_gray_background)
    worksheet.merge_range("N34:N37", "IEEE Recommended Rating (kV/μs)", center_format_gray_background)

    worksheet.merge_range("B38:B45", breaker_names, center_format)
    worksheet.merge_range("C38:C41", f"{local_station} (Local)", center_format)
    worksheet.merge_range("C42:C45", f"{remote_station} (Remote)", center_format)

    worksheet.write("D38", "10%", center_format)
    worksheet.write("D39", "30%", center_format)
    worksheet.write("D40", "60%", center_format)
    worksheet.write("D41", "100%", center_format)
    worksheet.write("D42", "10%", center_format)
    worksheet.write("D43", "30%", center_format)
    worksheet.write("D44", "60%", center_format)
    worksheet.write("D45", "100%", center_format)

    # Write output values
    write_small_table_output_values(worksheet, center_format, center_format_red_text, maximum_local_peaks_output, "10%", 38, IEEE_standards)
    write_small_table_output_values(worksheet, center_format, center_format_red_text, maximum_local_peaks_output, "30%", 39, IEEE_standards)
    write_small_table_output_values(worksheet, center_format, center_format_red_text, maximum_local_peaks_output, "60%", 40, IEEE_standards)
    write_small_table_output_values(worksheet, center_format, center_format_red_text, maximum_local_peaks_output, "100%", 41, IEEE_standards)
    write_small_table_output_values(worksheet, center_format, center_format_red_text, maximum_remote_peaks_output, "10%", 42, IEEE_standards)
    write_small_table_output_values(worksheet, center_format, center_format_red_text, maximum_remote_peaks_output, "30%", 43, IEEE_standards)
    write_small_table_output_values(worksheet, center_format, center_format_red_text, maximum_remote_peaks_output, "60%", 44, IEEE_standards)
    write_small_table_output_values(worksheet, center_format, center_format_red_text, maximum_remote_peaks_output, "100%", 45, IEEE_standards)

    workbook.close()

def write_output_values(worksheet, center_format, peaks_output, test_duty, row_1_number, row_2_number, row_3_number, IEEE_standards):
    worksheet.write(f"E{row_1_number}", peaks_output[test_duty]["kv"]["a"]["fault_type"], center_format)
    worksheet.write(f"F{row_1_number}", peaks_output[test_duty]["kv"]["a"]["fault_location"], center_format)
    worksheet.write(f"G{row_1_number}", peaks_output[test_duty]["kv"]["a"]["interrupting_current"], center_format)
    worksheet.write(f"H{row_1_number}", peaks_output[test_duty]["kv"]["a"]["peak_value"], center_format)

    worksheet.write(f"J{row_1_number}", peaks_output[test_duty]["rrrv"]["a"]["fault_type"], center_format)
    worksheet.write(f"K{row_1_number}", peaks_output[test_duty]["rrrv"]["a"]["fault_location"], center_format)
    worksheet.write(f"L{row_1_number}", peaks_output[test_duty]["rrrv"]["a"]["interrupting_current"], center_format)
    worksheet.write(f"M{row_1_number}", peaks_output[test_duty]["rrrv"]["a"]["peak_value"], center_format)

    worksheet.write(f"E{row_2_number}", peaks_output[test_duty]["kv"]["b"]["fault_type"], center_format)
    worksheet.write(f"F{row_2_number}", peaks_output[test_duty]["kv"]["b"]["fault_location"], center_format)
    worksheet.write(f"G{row_2_number}", peaks_output[test_duty]["kv"]["b"]["interrupting_current"], center_format)
    worksheet.write(f"H{row_2_number}", peaks_output[test_duty]["kv"]["b"]["peak_value"], center_format)

    worksheet.write(f"J{row_2_number}", peaks_output[test_duty]["rrrv"]["b"]["fault_type"], center_format)
    worksheet.write(f"K{row_2_number}", peaks_output[test_duty]["rrrv"]["b"]["fault_location"], center_format)
    worksheet.write(f"L{row_2_number}", peaks_output[test_duty]["rrrv"]["b"]["interrupting_current"], center_format)
    worksheet.write(f"M{row_2_number}", peaks_output[test_duty]["rrrv"]["b"]["peak_value"], center_format)

    worksheet.write(f"E{row_3_number}", peaks_output[test_duty]["kv"]["c"]["fault_type"], center_format)
    worksheet.write(f"F{row_3_number}", peaks_output[test_duty]["kv"]["c"]["fault_location"], center_format)
    worksheet.write(f"G{row_3_number}", peaks_output[test_duty]["kv"]["c"]["interrupting_current"], center_format)
    worksheet.write(f"H{row_3_number}", peaks_output[test_duty]["kv"]["c"]["peak_value"], center_format)

    worksheet.write(f"J{row_3_number}", peaks_output[test_duty]["rrrv"]["c"]["fault_type"], center_format)
    worksheet.write(f"K{row_3_number}", peaks_output[test_duty]["rrrv"]["c"]["fault_location"], center_format)
    worksheet.write(f"L{row_3_number}", peaks_output[test_duty]["rrrv"]["c"]["interrupting_current"], center_format)
    worksheet.write(f"M{row_3_number}", peaks_output[test_duty]["rrrv"]["c"]["peak_value"], center_format)

    worksheet.merge_range(f"I{row_1_number}:I{row_3_number}", IEEE_standards[test_duty]["TRV Peak"], center_format)
    worksheet.merge_range(f"N{row_1_number}:N{row_3_number}", IEEE_standards[test_duty]["RRRV"], center_format)

def write_small_table_output_values(worksheet, center_format, center_format_red_text, peaks_maximum_output, test_duty, row_1_number, IEEE_standards):
    worksheet.write(f"E{row_1_number}", peaks_maximum_output[test_duty]["kv"]["fault_type"], center_format)
    worksheet.write(f"F{row_1_number}", peaks_maximum_output[test_duty]["kv"]["fault_location"], center_format)
    interrupting_current_kv = peaks_maximum_output[test_duty]["kv"]["interrupting_current"]
    interrupting_current_kv = round(interrupting_current_kv, 2) if interrupting_current_kv != "N/A" else interrupting_current_kv
    worksheet.write(f"G{row_1_number}", interrupting_current_kv, center_format)
    peak_kv = peaks_maximum_output[test_duty]["kv"]["peak_value"]
    standards_peak_kv = IEEE_standards[test_duty]["TRV Peak"]
    if peak_kv == "N/A" or peak_kv < standards_peak_kv:
        peak_kv_format = center_format
    else:
        peak_kv_format = center_format_red_text
    rounded_kv = round(peak_kv) if peak_kv != "N/A" else peak_kv
    worksheet.write(f"H{row_1_number}", rounded_kv, peak_kv_format)

    worksheet.write(f"J{row_1_number}", peaks_maximum_output[test_duty]["rrrv"]["fault_type"], center_format)
    worksheet.write(f"K{row_1_number}", peaks_maximum_output[test_duty]["rrrv"]["fault_location"], center_format)
    interrupting_current_rrrv = peaks_maximum_output[test_duty]["rrrv"]["interrupting_current"]
    interrupting_current_rrrv = round(interrupting_current_rrrv, 2) if interrupting_current_rrrv != "N/A" else interrupting_current_rrrv
    worksheet.write(f"L{row_1_number}", interrupting_current_rrrv, center_format)
    peak_rrrv = peaks_maximum_output[test_duty]["rrrv"]["peak_value"]
    standards_peak_rrrv = IEEE_standards[test_duty]["RRRV"]
    if peak_rrrv == "N/A" or peak_rrrv < standards_peak_rrrv:
        peak_rrrv_format = center_format
    else:
        peak_rrrv_format = center_format_red_text
    rounded_rrrv = round(peak_rrrv) if peak_rrrv != "N/A" else peak_rrrv
    worksheet.write(f"M{row_1_number}", rounded_rrrv, peak_rrrv_format)

    worksheet.write(f"I{row_1_number}", standards_peak_kv, center_format)
    worksheet.write(f"N{row_1_number}", standards_peak_rrrv, center_format)
