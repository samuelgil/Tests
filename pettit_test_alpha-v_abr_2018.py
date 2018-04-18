# -*- coding: utf-8 -*-
"""
Created on Tue Aug 02 14:36:52 2016

@author: 71906377
By Samuel Gil for Integral SA
"""
import pandas as pd
import numpy as np
import Tkinter as tk
import tkFileDialog as filedialog
import xlwings as xw

root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilename()



def pettit (x, alpha):
    """   Input:
        y: vector of years - todavia no esta implementado 
        x: vector of data
        alpha: significance level (0.05 default)

    Output:
        
       
        K: tests statistics
        Ka:valor teorico de k en el nivel de probabilidad alpha
        Ho: inexistencia de cambio
        

    Examples
    --------
      >>> 
      >>>  
    
    x = np.random.rand(100)
    #genero rango de los datos     
    rang_x = [0] * len(x)
    for i, j in enumerate(sorted(range(len(x)), key=lambda y: input[y])):
        rang_x[j] = i
    """
    
 

    n=len(x)    
    alpha=0.05


#genero rango de los datos   
#==============================================================================
    def rango(x):
        """
        Genero el rango de los datos orden ascendente
        """
        rang_x = [0] * len(x)
        for i, j in enumerate(sorted(range(len(x)), key=lambda y: x[y])):
            rang_x[j] = i
    # forma mas eficiente de ordenar ni idea como funciona
    # a continuación empiezo la lista dede 1 hasta n         
        rango=[]
        for i in rang_x:
            j=i+1
            rango.append(j)
        return rango
    
    rango_x=rango(x)   
#==============================================================================
#genero lista de indices numero de dato
    n_dato=range(1,len(x)+1)

#calculo estadístico u
    u=[]
    l=0
    for j in n_dato:
        aux=rango_x[0:l+1]
        l=l+1
        #print aux
        uk=2*(sum(aux))-j*(n+1)
        u.append(uk)
    
#calculo valor absoluto de u
    abs_u=[abs(i) for i in u]    

#calculo estdístico K
    k=max(abs_u)

#calculo Ka
    ka=(((-np.log(alpha))*((n**3)+(n**2)))/6)**0.5
    
    if k>ka:
        Ho='Se Rechaza Ho'
    elif k<ka:
        Ho='Se Acepta Ho'
    else:
        Ho='Se Rechaza Ho'
    return k,ka,Ho
    
#==============================================================================
#ARCHIVO DE ORIGEN DE DATOS EN EXCEL

ruta=file_path
libro= xw.Book(ruta)
sht1=libro.sheets['Hoja1']

x=sht1.range('A1:P49').options(pd.DataFrame).value #escoger rango de datos 
#x_lista=x
#se remueven valores diferentes a float como NONE y caracteres especiales
       
#se remueven valores diferentes a float como NONE y caracteres especiales
x_lista=x.T.values.tolist()
maximos=[]
max_i=[]
for i in x_lista:
    max_i=[]
    for j in i:
       
        if type(j)==float:
            max_i.append(j)
    maximos.append(max_i)
#genero listas vacias para los resultados 
k_l=[]
ka_l=[]
Ho_l=[]
#

#calculos
#aplico la funcion mk_test
for i in maximos:
    k,ka,Ho = pettit(i,0.05)
    k_l.append(k)
    ka_l.append(ka)
    Ho_l.append(Ho)
    

#escribo resultados

sht1.range('A66').value='TEST DE PETTIT'
sht1.range('B66').value='Ho (Inexistencia de cambio)'
sht1.range('A67').value='ESTADISTICO K'
sht1.range('A68').value='K alfa'
sht1.range('A69').value='RESULTADO'

    
sht1.range('b67:p67').value=k_l
sht1.range('b68:p68').value=ka_l
sht1.range('b69:p69').value=Ho_l

    





