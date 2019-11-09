# -*- coding: utf-8 -*-
"""
Created on Sun Oct 27 08:21:32 2019

Functions for performing the random extraction!


@author: elena
"""
import random


# %% Function

def xmas_perm(J):
    '''
    Function for performing the permutation
    INPUT: J is a list of all the exclusions (every exclusion is a list)
            every people is mapped with a number, and it should exclude at least himself
    '''
    # init
    n = len(J)
    M = list(range(n))             # list of who_give index
    done = False                 # permutation done and correct!
    D = random.sample(M, len(M))  # 1st permutation (list of who_receive index)
    n_cycle = 0
    
    # permutation and check till result is OK
    while done == False:
        n_cycle += 1
        print('--> permut N '+str(n_cycle))
        
        # permutation
        D = random.sample(M, len(M))  # permutation (list of who_receive index)
        done = True
        
        # check exclusion
        for i in range(len(M)):
            if D[i] in J[i]:
                done = False    # permutation NO OK: redo!
    
    print('--> PERMUTATION OK!!!')
    return D



# %% Test 
if __name__ == '__main__':
    #       FUNCTION_PERMUTATION WITH CONSTRAINTS IN INPUT
    # We generate an ordered sequence of n variables in form of a list M
    # The numbers correspond to the people
    # Each guy cannot give away a gift to some specific guys: this request is expressed by list J
    # The output is a permutation D
    # Everyone M[i] have to prepare a gist to D[i]
    
    # J[i] = people excluded in the realization of the gift prepared by i
    # in J[i] there is by default i and his/her partner
    J = [[0],[1,2],[2,1],[3,4],[4,3],[5,6],[6,5],[7,8],[8,7],[9],[10,11,12],[10,11,12],[10,12,11],[13,14],[14,13],[15,16],[16,15]]
    
    # permuted list
    D = xmas_perm(J)
    

