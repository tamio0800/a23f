# Functions Returning Boolean Value
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
            
# Functions Counting Certain Numbers
    1. a23f_num_unique(target_rng as range, optional exclude_empty=1)
        This function will return how many different values are in a range.
    
    2. a23f_num_isin_cell(target_char, target_cell, Optional fuzzy = 0)
        Returns how many target_chars are in the target_cell.
        
    3. a23f_num_isin_range(target_cell, target_rng As Range, Optional fuzzy = 0)
        Returns how many string/characters in target_cell are in the target_rng.
        
# Functions Getting Certain Values From A Cell Or A Range
    1. a23f_get_after(target_string, separator)
        Returns the part in target_string after the first separator in target_string, 
        returns whole target_string while none of separator was in it.
    
    2. a23f_get_before(target_string, separator)
        Returns the part in target_string before the first separator in target_string,
        returns whole target_string while none of separator was in it.

    3. a23f_get_in_between(target_cell, first_char, second_char, Optional separator = ", ")
        Returns the part of target_cell between first_char and second_char (both excluded),
        and if there were more than one pairs of first_char and second_char, 
        it will return all texts in them and separate them by separator.
        a23f_get_in_between("Hi (John) and (Mary).", "(", ")"): "John, Mary"
        
    4. a23f_get_nth_after_split(target_string, nth, split_by)
        a23f_get_nth_after_split("1st| 2nd| 3rd| 4th| 5th| 6th", 3, "|"): "3rd" (with trim() function inside)
        
    5. a23f_get_nth_most_frequent(target_rng As Range, nth, Optional return_id = 1, Optional separator = ", ")
        Say there is a list you have to sign your name up every time you go into gym, it looks like below: 
        Name
        Tamio
        Michael
        Nilson
        Jack
        Tamio
        Tom
        Jack
        Tamio
        ...
        And you want to find out who goes to gym most frequently, just apply this:
        a23f_get_nth_most_frequent(the_name_column, 1) -> "Tamio",
        and the frequencies are a23f_get_nth_most_frequent(the_name_column, 1, 0) -> 3.
        This function is quite useful when you want to simply claculate a neat and small frequency table.       
        
# Functions Doing Calculations
    1. a23f_calc_percentage(target_cell, target_rng As Range, Optional exponent = 1)
        Returns target_cell^(exponent) / SIGMA[target_rng.each_value^(exponent)],
        when expoent equals 1, this function returns the portion of target_cell's value among target_rng.
        
    2. 23f_calc_reverse_percentage(target_cell, target_rng As Range, Optional exponent = 1)
        Say there is a series of number: 10, 20, 30, 40, 50 called A_Series,
        23f_calc_reverse_percentage(10, A_Series, 2) =
            (10^(-1))^(2) / SIGMA[(target_rng.each_value^(-1))^(2)] ≒ 0.683242
        
# Functions Doing Transformations
    1. a23f_to_datetime(target_cell, Optional return_datevalue = 1)
        Returns date value of str(date like value),
        could apply on 'yyyy-mm-dd ...', 'yyyy/mm/dd ...', 'mm-dd-yyyy ...' ...and so on.
        
    
    
    
    
    
        

