# Functions Return Boolean Value
    1. a23f_all_the_same(target_rng as range, Optional exclude_empty = 1)
        target_rng -> the Excel range you want to check if the values are all the same.
        exclude_empty -> if it was set to 0, this function will take empty cell as a value.
        a23f_all_the_same -> return True if all values in the range are all the same, False if not.
    
    2. a23f_all_the_different(target_rng as range, Optional exclude_empty = 1)
        target_rng -> the Excel range you want to check if the values are all different.
        exclude_empty -> if it was set to 0, this function will take empty cell as a value.
        a23f_all_the_different -> return True if all values in the range are all different, False if not. 
    
    3. a23f_isin_cell(target_char, target_cell, Optional fuzzy = 0)
        target_char -> the character or a string you want to check if it was in a Excel cell.
        target_cell -> the Excel cell you want to find if a certain character or string was in it.
        fuzzy -> 
            a23f_isin_cell("hi", "Hi There!", 0): False
            a23f_isin_cell("hi", "Hi There!", 1): True
    
    4. a23f_isin_range(target_cell, target_rng As Range, Optional fuzzy = 0)
        target_cell -> the Excel cell or a string you want to find if it was in Excel ranges.
        target_rng -> the Excel ranges you want to find if it contains the certain cell.value or a string.
        fuzzy ->
            a23f_isin_range("hi", ["Hello!", "Haruhi Lyn"], 0): False
            a23f_isin_range("hi", ["Hello!", "Haruhi Lyn"], 0): True
