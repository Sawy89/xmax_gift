# -*- coding: utf-8 -*-
"""
Created on Sun Oct 27 08:21:32 2019

@author: elena
"""
import random

from permutation import toname # to invert the number to the associated name

#       FUNCTION_PERMUTATION WITH CONSTRAINTS IN INPUT
# We generate an ordered sequence of n variables in form of a list M
# The numbers correspond to the people
# Each guy cannot give away a gift to some specific guys: this request is expressed by list J
# The output is a permutation D
# Everyone M[i] have to prepare a gist to D[i]

def xmas_perm(J, n):
    # lista dei Mittenti
    M = list(range(n))
    
    # lista dei Destinatari
    D = random.sample(M, len(M))
    
    # Exclude people  
    t = len(D)-1 
    while t >= 0:    
        for j in range(n):
            if D[j] in J[j]:
               random.shuffle(D) 
        t -= 1
    return D


 
if __name__ == '__main__':
    
    # J[i] = people excluded in the realization of the gift prepared by i
    # in J[i] there is by default i and his/her partner
    J = [[0],[1,2],[2,1],[3,4],[4,3],[5,6],[6,5],[7,8],[8,7],[9],[10,11,12],[10,11,12],[10,12,11],[13,14],[14,13],[15,16],[16,15]]
    
    # permuted list
    D = xmas_perm(J, 17)
    
    # Now we check with the names (toname) if there are some wrong exclusions
    
    # Creo coppi nomi MITTENTE-DESTINATARIO
    # lista dei Mittenti
    M = list(range(17))
    M = toname(M)
    D = toname(D)
    
    # stampo le coppie
    print('(Mittente , Destinatario)\n')
    t = [(M[i]+ ' fa il regalo a ' + D[i]) for i in range(17)]     
    print(t)
  


