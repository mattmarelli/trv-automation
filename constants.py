test_duty_table = {
    "245": {
        "10%": {
            "TRV Peak": 459,
            "RRRV": 7,
        },
        "30%": {
            "TRV Peak": 400,
            "RRRV": 5,
        },
        "60%": {
            "TRV Peak": 390,
            "RRRV": 3,
        },
        "100%": {
            "TRV Peak": 364,
            "RRRV": 2,
        }
    },
    "550": {
        "10%": {
            "TRV Peak": 1030,
            "RRRV": 7,
        },
        "30%": {
            "TRV Peak": 899,
            "RRRV": 5,
        },
        "60%": {
            "TRV Peak": 876,
            "RRRV": 3,
        },
        "100%": {
            "TRV Peak": 817,
            "RRRV": 2,
        }
    },
}

fault_types_table = {
    1: "A-G",
    2: "B-G",
    3: "C-G",
    4: "AB-G",
    5: "AC-G",
    6: "BC-G",
    7: "ABC-G",
    8: "AB",
    9: "AC",
    10: "BC",
    11: "ABC",
}

fault_locations_table = {
    1: "Local Close-In",
    2: "Local Short-Line",
    3: "Mid-Line",
    4: "Remote Short-Line",
    5: "Remote Close-In",
}