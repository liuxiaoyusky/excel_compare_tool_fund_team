from seg_mapping_config import sec_mapping

bond_code_set = {
    'MR',
    'MF',
    'G',
    'CB',
    'B',
}

stack_code_set = {
    'S',
    'P',
}


missing_isin_or_stack_code_mapping_dict = sec_mapping

# 绝对容差：当 |hsbc - spectra| <= 此值时视为相等
TOLERANCE_ABS = 0.000