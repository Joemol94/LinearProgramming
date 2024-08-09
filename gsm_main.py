import pandas as pd
import itertools
# data_cpy = pd.read_excel("data/gsm_data_dummy.xlsx")
# tl_white = [127,175,300]
# flutting =  [127,150,125,100,165,140,152,143,115]
# brown_l2_l3 = [127,175,125,150,140,170,149,165,500,790,795,809,775,115] 

def find_best_match(data_cpy,tl_white,flutting, brown_l2_l3):
    best_a = None
    best_b = None
    best_c = None
    best_sum = None
    best_diff = float('inf')
    for target_gsm in data_cpy['target_gsm'].unique():
        subset = data_cpy[data_cpy['target_gsm'] == target_gsm]
        if (subset['topliner'] == 'BROWN').any() and (subset['ply']==3).any():
            for a, b, c in itertools.product(brown_l2_l3, flutting, brown_l2_l3):
                expression_value = sum([a + (1.43 * b) + c])
                diff = abs(expression_value - target_gsm)
                if diff < best_diff:
                        best_a = a
                        best_b = b
                        best_c = c
                        best_diff = diff
                        best_sum = expression_value 
            values_dict = {'tl_l1': int(best_a), 'tl_f1': int(best_b), 'tl_l2': int(best_c),'total_optimized_gsm':int(best_sum)}
            mask = (data_cpy['target_gsm'] == target_gsm)
            data_cpy.loc[mask, ['tl_l1', 'tl_f1', 'tl_l2','total_optimized_gsm']] = [values_dict['tl_l1'], values_dict['tl_f1'], values_dict['tl_l2'], values_dict['total_optimized_gsm']]
            best_a = None
            best_b = None
            best_c = None
            best_sum = None
            best_diff = float('inf')
            print(values_dict)

        elif (subset['topliner'] == 'WHITE').any() and (subset['ply']==3).any():
            for a, b, c in itertools.product(tl_white, flutting, brown_l2_l3):
                expression_value = sum([a + (1.43 * b) + c])
                diff = abs(expression_value - target_gsm)
                if diff < best_diff:
                        best_a = a
                        best_b = b
                        best_c = c
                        best_diff = diff
                        best_sum = expression_value 
            values_dict = {'tl_l1': int(best_a), 'tl_f1': int(best_b), 'tl_l2': int(best_c),'total_optimized_gsm':int(best_sum)}
            mask = (data_cpy['target_gsm'] == target_gsm)
            data_cpy.loc[mask, ['tl_l1', 'tl_f1', 'tl_l2','total_optimized_gsm']] = [values_dict['tl_l1'], values_dict['tl_f1'], values_dict['tl_l2'], values_dict['total_optimized_gsm']]
            print(values_dict)
            best_a = None
            best_b = None
            best_c = None
            best_sum = None
            best_diff = float('inf')

        elif (subset['topliner'] == 'BROWN').any() and (subset['ply']==5).any():
            for a, b, c in itertools.product(brown_l2_l3, flutting, brown_l2_l3):
                expression_value = sum([(2* 1.43 * b) + (3*c)])
                diff = abs(expression_value - target_gsm)
                if diff < best_diff:
                        best_a = a
                        best_b = b
                        best_c = c
                        best_diff = diff
                        best_sum = expression_value 
            values_dict = {'tl_l1': int(best_a), 'tl_f1': int(best_b), 'tl_l2': int(best_c),'tl_f2': int(best_b), 'tl_l3': int(best_c),'total_optimized_gsm':int(best_sum)}
            print(values_dict)
            mask = (data_cpy['target_gsm'] == target_gsm)
            data_cpy.loc[mask, ['tl_l1', 'tl_f1', 'tl_l2','tl_f2','tl_l3','total_optimized_gsm']] = [values_dict['tl_l1'], values_dict['tl_f1'], values_dict['tl_l2'], values_dict['tl_f2'], values_dict['tl_l3'], values_dict['total_optimized_gsm']]
            best_a = None
            best_b = None
            best_c = None
            best_sum = None
            best_diff = float('inf')        
        elif (subset['topliner'] == 'WHITE').any() and (subset['ply']==5).any():
            for a, b, c in itertools.product(tl_white, flutting, brown_l2_l3):
                expression_value = sum([(2* 1.43 * b) + (3*c)])
                diff = abs(expression_value - target_gsm)
                if diff < best_diff:
                        best_a = a
                        best_b = b
                        best_c = c
                        best_diff = diff
                        best_sum = expression_value 
            values_dict = {'tl_l1': int(best_a), 'tl_f1': int(best_b), 'tl_l2': int(best_c),'tl_f2': int(best_b), 'tl_l3': int(best_c),'total_optimized_gsm':int(best_sum)}
            print(values_dict)
            mask = (data_cpy['target_gsm'] == target_gsm)
            data_cpy.loc[mask, ['tl_l1', 'tl_f1', 'tl_l2','tl_f2','tl_l3','total_optimized_gsm']] = [values_dict['tl_l1'], values_dict['tl_f1'], values_dict['tl_l2'], values_dict['tl_f2'], values_dict['tl_l3'], values_dict['total_optimized_gsm']]
            best_a = None
            best_b = None
            best_c = None
            best_sum = None
            best_diff = float('inf')
        else:
            print('else')

    
    return data_cpy
    
# opt_data = find_best_match(data_cpy,tl_white,flutting, brown_l2_l3)
# print(opt_data)

