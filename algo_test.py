#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Jun 21 04:40:56 2019
The progam is tested and working correctly if incase of any issues. please contact jnikhil@seas.upenn.edu
Pakages and dependancies to be install before the program runs are pandas, numpy, itertools 
Pandas data frame properties are primarily used for large data operations

Two different definations are given for test statistic on Wikipedia and other source. The program function returns W statistic 
based on both methods.

@author: jnikhil
"""

import pandas as pd
import numpy as np
from pandas import ExcelWriter
from pandas import ExcelFile
import itertools as it

def wilcoxon_signed_rank_test(data_df):
    
    '''
    The function takes pandas data frame and returns the dictionary of pair of products
    and itsWilcoxon Signed Rank Test statistic. 
    **It returns two different dictionaries based on defination for Test statistic on Wikipedia and other university source
    Both W statistics represnets same major 
    '''
    frame_cols = data_df.columns
    required_cols = frame_cols[2:]
    dict_pair_wiki = {}
    dict_pair_theory = {}
    
    for i in range(len(required_cols)):
        wtest_df = pd.DataFrame(columns = ['key', 'nonkey', 'abs_diff','sgn','rank','sgn_x_rank'])
        wtest_df['key'] = data_df.iloc[:,1]
        wtest_df['nonkey'] = data_df.iloc[:,i+2]
        wtest_df['abs_diff'] = abs(wtest_df['key'] - wtest_df['nonkey'])
        wtest_df['sgn'] = np.sign(wtest_df['key'] - wtest_df['nonkey'])
        wtest_df = wtest_df[(wtest_df[['abs_diff']] != 0).all(axis=1)]
        wtest_df = wtest_df.sort_values(by=['abs_diff'])
        wtest_df['rank'] = wtest_df['abs_diff'].rank(ascending=1) #ranked from lowest to highest, with tied ranks included where appropriate
        wtest_df['sgn_x_rank'] = wtest_df['sgn']*wtest_df['rank']
        negative_df = wtest_df[(wtest_df[['sgn_x_rank']] < 0).all(axis=1)]
        positive_df = wtest_df[(wtest_df[['sgn_x_rank']] > 0).all(axis=1)]
        nrank_sum = abs(negative_df['sgn_x_rank'].sum())
        prank_sum = positive_df['sgn_x_rank'].sum() 
        
        # Wikipedia - W stat is equal to sum of all singed ranks 
        dict_pair_wiki.update({(frame_cols[1],frame_cols[i+2]):wtest_df['sgn_x_rank'].sum()})
        
        # other source - W stat is equal to min of sum of negtive ranks and sum of positive ranks
        dict_pair_theory.update({(frame_cols[1],frame_cols[i+2]):min(nrank_sum,prank_sum)})
    
    return (dict_pair_wiki,dict_pair_theory) 


def permu_n_weights(data_df):
    
    '''
    This function calculates permu with weights i.e combinations of products to be compared and returns new data frame. 
    Using 100% equal weighting tp calculate the weighted combination of the columns of products to be compared .The function also returns
    list of combinations along with new data frame. 
    '''
    
    all_columns = data_df.columns
    nonkey_cols = all_columns[2:]
    prod_combinations = []
    
    for i in range(len(nonkey_cols)-1):
        prod_combinations.extend(list(it.combinations(nonkey_cols, i+2)))  
    
    for element in prod_combinations:
        sum = 0
        col_name = "W_"
        for sub in element:
            S = sub
            sum = sum+data_df[S]
            col_name = col_name + sub + "_"
        col_name = col_name[:-1]
        data_df[col_name] = sum/(len(element))
    
    PerW_df = data_df.drop(nonkey_cols, axis=1) 
   
    return (prod_combinations,PerW_df)


def main():
    
    '''
    The data file is always assumed to be having first column as identifier, second as a 'main Product'
    and remaining columns are products to be compared
    '''
    
    data_df = pd.read_excel('Algodata.xlsx', sheet_name='Data')
    data_df = data_df.drop_duplicates(subset='Customer', keep="last")
    # Removes the duplicate customers if any 
    
    dict_pair_wiki,dict_pair_theory = wilcoxon_signed_rank_test(data_df)
    
    print("\nKey and Non Key product pair with W test (Wikipedia Defination):\n")
    #print(dict_pair_wiki)
    for index, (key, value) in enumerate(dict_pair_wiki.items()):
        print( str(index) + " : " + str(key) + " : " + str(value) )
    
    print("\nKey and Non Key product pair with W test (Book Defination):\n")
    #print(dict_pair_theory)
    for index, (key, value) in enumerate(dict_pair_theory.items()):
        print( str(index) + " : " + str(key) + " : " + str(value) )
    
    prod_comb,PerW_df= permu_n_weights(data_df)
    pw_pair_wiki,pw_pair_theory = wilcoxon_signed_rank_test(PerW_df)
     
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