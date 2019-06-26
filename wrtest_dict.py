#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Jun 22 08:21:40 2019
The progam is tested and working correctly if incase of any issues. please contact jnikhil@seas.upenn.edu
Two different definations are given for test statistic on Wikipedia and other source. The program function returns W statistic 
based on both methods. 
Two ranking methods are written however efficient on is used. Another is simple but not efficient O(N^2)

@author: jnikhil
"""

import itertools as it
import xlrd 
import copy

def average_rank(w_list):    
    rank_sum = 0
    dcount = 0
    rank_list = [0 for x in range(len(w_list))] 
    list_x = sorted(range(len(w_list)),key=w_list.__getitem__)
    list_y = [w_list[pos] for pos in list_x]    
    
    for i in range(len(w_list)):
        rank_sum = rank_sum+i
        dcount = dcount+1
        if i==len(w_list)-1 or list_y[i] != list_y[i+1]:
            for j in range(i-dcount+1,i+1):
                rank_avg = rank_sum / float(dcount) + 1
                rank_list[list_x[j]] = rank_avg
            rank_sum = 0
            dcount = 0 
            
    return rank_list

'''
#simple but O(N^2) complexity and less efficent method 
def average_rank(w_list):
    rank_list=[]
    for ele_1 in range(len(w_list)):
            odr=1
            du=1
            for ele_2 in range(len(w_list)):
                if ele_2 != ele_1 and w_list[ele_2] < w_list[ele_1]: 
                    odr = odr + 1
                if ele_2 != ele_1 and w_list[ele_2] == w_list[ele_1]: 
                     du = du + 1     
            rank_list.append(odr + (du - 1) / 2)
    return rank_list
'''

def wilcoxon_signed_rank_test(data_dict):
    
    '''
    The function takes data and returns the dictionary of pair of products
    and itsWilcoxon Signed Rank Test statistic. 
    **It returns two different dictionaries based on defination for Test statistic on Wikipedia and other university source
    Both W statistics represnets same major 
    '''
    data_cols = list(data_dict.keys())
    required_cols = data_cols[2:]
    dict_pair_wiki = {}
    dict_pair_theory = {}
    
    for i in range(len(required_cols)):
        rank_list = []
        wtest_dict = {}
        wtest_dict['key'] = data_dict[data_cols[1]]
        wtest_dict['nonkey'] = data_dict[data_cols[i+2]]
        
        wtest_dict['diff'] = [((wtest_dict['key'][i]-wtest_dict['nonkey'][i])) for i in range(len(wtest_dict['key'])) if ((wtest_dict['key'][i]-wtest_dict['nonkey'][i]) !=0)]

        wtest_dict['abs_diff'] = [(abs(wtest_dict['diff'][i])) for i in range(len(wtest_dict['diff']))]

        wtest_dict['sgn'] = [1 if (wtest_dict['diff'][i])>0 else -1 for i in range(len(wtest_dict['diff']))]
        
        rank_list = average_rank(wtest_dict['abs_diff'])
        
        wtest_dict['rank'] = rank_list
        #ranked from lowest to highest, with tied ranks included where appropriate
        
        wtest_dict['sgn_x_rank'] =  [(wtest_dict['sgn'][i])*(wtest_dict['rank'][i]) for i in range(len(wtest_dict['diff']))]  
        
        dict_pair_wiki.update({(data_cols[1],data_cols[i+2]):sum(wtest_dict['sgn_x_rank'])})
        
        wtest_dict['neg_rank'] = [(wtest_dict['sgn_x_rank'][i]) for i in range(len(wtest_dict['diff'])) if (wtest_dict['sgn_x_rank'][i]) <0 ]
        
        wtest_dict['pos_rank'] = [(wtest_dict['sgn_x_rank'][i]) for i in range(len(wtest_dict['diff'])) if (wtest_dict['sgn_x_rank'][i]) >0 ]  
        
        nrank_sum = abs(sum(wtest_dict['neg_rank']))
        prank_sum = sum(wtest_dict['pos_rank'])
       
        # Wikipedia - W stat is equal to sum of all singed ranks 
        dict_pair_wiki.update({(data_cols[1],data_cols[i+2]):sum(wtest_dict['sgn_x_rank'])})
        
        # other source - W stat is equal to min of sum of negtive ranks and sum of positive ranks
        dict_pair_theory.update({(data_cols[1],data_cols[i+2]):min(nrank_sum,prank_sum)})
       
    
    return (dict_pair_wiki,dict_pair_theory) 


def permu_n_weights(data_dict):
    
    '''
    This function calculates permu with weights i.e combinations of products to be compared and returns new data frame. 
    Using 100% equal weighting tp calculate the weighted combination of the columns of products to be compared . The function also returns
    list of combinations along with new dictionary. 
    '''    
    comb_dict={}
    comb_dict=copy.deepcopy((data_dict))
    all_columns = list(comb_dict.keys())
    nonkey_cols = all_columns[2:]
    prod_combinations = []
    
    for i in range(len(nonkey_cols)-1):
        prod_combinations.extend(list(it.combinations(nonkey_cols, i+2)))  
    
    for element in prod_combinations:
        sum=[0 for x in range(len(data_dict[all_columns[1]]))] 
        col_name = "W_"
        for sub in element:
            S = sub
            sum=[x+y for x,y in zip(sum, comb_dict[S])]
            col_name = col_name + sub + "_"
        col_name = col_name[:-1]
        comb_dict[col_name] = [ val/len(element)  for val in sum] 
    
    for ele in nonkey_cols:
        del comb_dict[ele] 
        
    return (prod_combinations,comb_dict)   


def main():
    
    '''
    The data file is always assumed to be having first column as identifier, second as a 'main Product'
    and remaining columns are products to be compared
    '''  
    loc = ("yourdata.xlsx")               # Change the file name here 
    wb = xlrd.open_workbook(loc) 
    sheet = wb.sheet_by_index(1)          # Change the sheet number here 
    data_dict = {}
    pos=0
    col_names = sheet.row_values(0)
    for element in sheet.row_values(0):
        data_dict.update({element:sheet.col_values(pos)[1:]})
        pos += 1
    
    
    dict_pair_wiki,dict_pair_theory = wilcoxon_signed_rank_test(data_dict)
    
    print("\nKey and Non Key product pair with W test (Wikipedia Defination):\n")
    #print(dict_pair_wiki)
    for index, (key, value) in enumerate(dict_pair_wiki.items()):
        print( str(index) + " : " + str(key) + " : " + str(value) )
    
    print("\nKey and Non Key product pair with W test (Book Defination):\n")
    #print(dict_pair_theory)
    for index, (key, value) in enumerate(dict_pair_theory.items()):
        print( str(index) + " : " + str(key) + " : " + str(value) )
    
    prod_comb,PerW_dict = permu_n_weights(data_dict)
    pw_pair_wiki,pw_pair_theory = wilcoxon_signed_rank_test(PerW_dict)
     
    print("\nNumber of Product Combinations:\n")
    print(len(prod_comb))
     
    print("\nProduct Combinations:\n")
    print((prod_comb))
    
    print("\nKey and Non key combination with W test (Wikipedia Defination):\n")
    #print(pw_pair_wiki)
    for index, (key, value) in enumerate(pw_pair_wiki.items()):
        print( str(index) + " : " + str(key) + " : " + str(value) )
    
    print("\nKey and Non key combination with W test (Book Defination):\n")
    #print(pw_pair_theory)
    for index, (key, value) in enumerate(pw_pair_theory.items()):
        print( str(index) + " : " + str(key) + " : " + str(value) )
    
    bestfit_nonkey = min(dict_pair_wiki, key=lambda y: abs(dict_pair_wiki[y]))
    print("\nBest Fit non key product:\n")
    print(bestfit_nonkey)     #Best fit non key product for which pair has least absolute W stat 
                              # The value is near to zero or postive and negative are approx.equal 
    
    print("\nBest Fit non key product combinations:\n")
    bestfit_pernw = min(pw_pair_wiki, key=lambda y: abs(pw_pair_wiki[y]))
    print(bestfit_pernw)      # Best fit combination for which test statisitc is near to zero or least among all
    
if __name__ == "__main__":
    main()