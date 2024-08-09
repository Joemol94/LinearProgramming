import pandas as pd
import numpy as np
from pulp import *
from datetime import datetime


#var_lst,opt_val = [],[]
def process_df(final_df,trim_val):
    df = final_df.copy()
    df['carton_Demand'] = df['demand'] #monthly_demand - june
    df['Demand_conv'] = df['carton_Demand']/100 ####do not divide 100
    df['Demand'] = df['Demand_conv'] * df['width_type']
    reel_size = 2200-trim_val
    n = len(df)
    sku_name_list = list(df['sku'])
    final_output_df = pd.DataFrame(columns=[f"{i}" for i in sku_name_list])
    return df,n,reel_size,final_output_df

def initial_model(df,n,reel_size,final_output_df):
    # global var_lst
    # global opt_val
    var_lst,opt_val = [],[]
    #y = [LpVariable(name=f"y{i}", lowBound=0, upBound=100,cat='Integer') for i in range(n)]
    y = [LpVariable(name=f"y{i}", lowBound=0, upBound=5,cat='Integer') for i in range(n)]
    model0 = LpProblem(name='OptProblem', sense=LpMaximize)
    model0 += sum(df.width_type[i] * y[i] for i in range(n)), "objective function"
    for i in range(n):
        model0 += y[i] * df.width_type[i] <= df.Demand[i] , f"DemandReq{i + 1}"
    model0 += sum(df.width_type[i] * y[i] for i in range(n))<=reel_size,'mainconstraint'
    model0 += sum(y)<=5 ,"cutsconstraint"
    model0.solve()
    for v in y:
        var_lst.append(v.varValue)
    final_output_df.loc[len(final_output_df)] = var_lst
    opt_val.append(value(model0.objective))
    return final_output_df,var_lst,opt_val

def run_model(final_output_df,df,n,reel_size,var_lst,opt_val):
    #x = [LpVariable(name=f"x{i}", lowBound=0, upBound=100,cat='Integer') for i in range(n)]
    x = [LpVariable(name=f"x{i}", lowBound=0, upBound=5,cat='Integer') for i in range(n)]
    model = LpProblem(name='OptProblem', sense=LpMaximize)
    model += sum(df.width_type[i] * x[i] for i in range(n)), "objective function"
    for k in range(n):
        model += x[k] * df.width_type[k] <= (df.Demand[k]-(var_lst[k]* df.width_type[k])) , f"DemandReq{k + 1}"
    model += sum(df.width_type[j] * x[j] for j in range(n))<=reel_size,'mainconstraint' 
    model += sum(x)<=5 ,"cutsconstraint"  
    model.solve()
    for l in range(n):
        df.Demand[l] = df.Demand[l]-(var_lst[l]* df.width_type[l])    
    for v in model.variables():
        var_lst.append(v.varValue)
    opt_val.append(value(model.objective))    
    del model.constraints
    del var_lst[0:n]
    final_output_df.loc[len(final_output_df)] = var_lst           
    return final_output_df
def check_demand_const(df):
    if (df['Demand'] == 0).all() or (df['Demand']<0).any() or (df['Demand']<df['width_type']).all():
        return False
    else:
        return True
def run_model_terminate(final_output_df,df,n,reel_size,var_lst,opt_val):
    check_demand = check_demand_const(df)
    while check_demand:
        run_model(final_output_df,df,n,reel_size,var_lst,opt_val)
        check_demand = check_demand_const(df)
    return final_output_df,var_lst,opt_val

def generate_output_df(final_output_df,df,var_lst,opt_val):
    # global var_lst
    # global opt_val
    final_output_df['opt_val'] = opt_val
    del var_lst
    del opt_val
    final_output_df = final_output_df.drop(final_output_df.index[len(final_output_df)-1])
    print('before resetting: ')
    print(final_output_df)
    #final_output_df_cpy = final_output_df.groupby(list(final_output_df.columns)).size().reset_index(name='Iteration')
    final_output_df_cpy = final_output_df.groupby(list(final_output_df.columns)).size().reset_index(name='Iteration')
    final_output_df_cpy['Iteration'] = final_output_df_cpy['Iteration']*100
    print('after column reset: ')
    print(final_output_df_cpy)
    output_df = final_output_df_cpy.copy()
    print('what col?')
    print(output_df.iloc[:, :-2].sum(axis=1))
    #output_df['total_demand'] = output_df['Iteration'] * 100 * output_df.iloc[:, :-2].sum(axis=1) ####do not multiply by 100
    output_df['total_demand'] = output_df['Iteration'] * output_df.iloc[:, :-2].sum(axis=1)
    output_df = output_df.rename(columns={'opt_val':'Optimum Reel Size (in mm)','total_demand':'Total Demand (in cartons)'})

    return output_df

def main_process(df_1,trim_val):
    grouped_df = df_1.groupby(['length', 'gsm'])
    output_df_final = pd.DataFrame()
    width_value_list,length_value_list,gsm_value_list = [],[],[]
    for group_name, group_df in grouped_df:
        group_df = group_df.reset_index()
        width_value_list.append(list(group_df['width_type']))
        length_value_list.append(list(group_df['length']))
        gsm_value_list.append(list(group_df['gsm']))
        df,n,reel_size,final_output_df = process_df(group_df,trim_val)
        final_output_df,var_lst,opt_val = initial_model(df,n,reel_size,final_output_df)
        final_output_df = run_model(final_output_df,df,n,reel_size,var_lst,opt_val)
        final_output_df,var_lst,opt_val = run_model_terminate(final_output_df,df,n,reel_size,var_lst,opt_val)
        output_df = generate_output_df(final_output_df,df,var_lst,opt_val)
        output_df_final = pd.concat([output_df_final, output_df], ignore_index=True) 
    return output_df_final,width_value_list,length_value_list,gsm_value_list

def preprocess_output_data(output_df_final,width_value_list,length_value_list,gsm_value_list):
    output_df_final = output_df_final.fillna(0)
    cols = list(output_df_final.columns.values) 
    cols.pop(cols.index('Optimum Reel Size (in mm)'))
    cols.pop(cols.index('Iteration'))
    cols.pop(cols.index('Total Demand (in cartons)'))
    output_df_final = output_df_final[cols+['Optimum Reel Size (in mm)','Iteration','Total Demand (in cartons)']]
    output_df_final = output_df_final.sort_values(by='Optimum Reel Size (in mm)', ascending=False)
    old_df = pd.DataFrame(output_df_final)
    new_data = [[item for sublist in width_value_list for item in sublist],
                [item for sublist in length_value_list for item in sublist],
                [item for sublist in gsm_value_list for item in sublist]]
    new_df = pd.DataFrame(new_data, columns=old_df.columns[:-3])
    output_df_final_mod = pd.concat([new_df, old_df], axis=0,ignore_index=True)
    output_df_final_mod = output_df_final_mod.reset_index(drop=True)
    first_three_index_col = ['Width','Length','GSM']
    output_df_final_mod.index = [f"{i}" for i in first_three_index_col] + [f"Demand Allocation {i-3}" for i in range(4,len(output_df_final_mod)+1)]
    output_df_final_mod = output_df_final_mod.fillna('-')

    del width_value_list
    del length_value_list
    del gsm_value_list
    return output_df_final_mod

def paper_req_calc(output_df_final_mod):
    df = output_df_final_mod.copy()
    paper_req_ton_list = []
    for index_num in range(3,len(df)):
        length_values = df.iloc[df.index.get_loc('Length'), :][df.iloc[index_num] > 0][0]
        length_in_m = length_values/1000
        gsm_values = df.iloc[df.index.get_loc('GSM'), :][df.iloc[index_num] > 0][0]
        reel_size_values = df.iloc[index_num,-3]
        reel_size_in_m = reel_size_values/1000
        iteration_values = df.iloc[index_num,-2]
        paper_req_gm_list = length_in_m * gsm_values*reel_size_in_m*iteration_values
        paper_req_ton_list.append(paper_req_gm_list/1000000)

    rounded_list = [round(x, 6) for x in paper_req_ton_list]
    df['Paper Requirement (in tons)'] = ['-','-','-']+rounded_list
    
    df['Index'] = df.index
    df = df[['Index'] + df.columns[:-1].tolist()]
    del paper_req_ton_list
    return df
#[len (in m)*reel_width(in m)*gsm(gm/m^2)*no_of iteration]/10^6

# changes
####do not 
#### no of cuts should not exceed 5