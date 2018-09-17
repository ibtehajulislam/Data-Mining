# Importing packages:
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import xlrd
import csv
import pylab
from sklearn.cluster import KMeans


## Get data from Excel file
def read_data (filename):
    workbook = xlrd.open_workbook(filename)
    print('Reading Excel Data sheets')

    while True:
        names = workbook.sheet_names()
        print ("Please select a sheet from below name list : ")
        print (names)
        sheet_name = input('Please Enter the Sheet Name: ').strip()
        worksheet = workbook.sheet_names()
        if sheet_name in worksheet:
            n_worksheet = workbook.sheet_by_name(sheet_name)
            print ('The sheet is Available')

            #Reading evey Row and Column
            
            total_rows = n_worksheet.nrows
            total_cols = n_worksheet.ncols
            table = list()
            record = list()
            for x in range (1,total_rows):
                for y in range (total_cols):
                    record.append(n_worksheet.cell(x,y).value)
                table.append(record)
                record = []
                x += 1
            table = np.array(table)
            #print(table)
            if n_worksheet == workbook.sheet_by_name(sheet_name):
                break
        else:
            print('The sheet is NOT Available. Please select from the below list only.')
    
    return table


def run_KMean(number_of_class,dataX):

    cmap = plt.cm.get_cmap('hsv', number_of_class+1) # for color map on Plot

    kmeans = KMeans(n_clusters = number_of_class, init = 'k-means++', max_iter = 300, n_init = 10, random_state = 0) 
    y_kmeans = kmeans.fit_predict(dataX) # Creating the model on dataset

    for i in range(0,number_of_class): # for plotting the dataset
        plt.scatter(dataX[y_kmeans == i, 0], dataX[y_kmeans == i, 1], s = 5, c = cmap(i+1), label = 'cluster_'+str(i+1))

    plt.scatter(kmeans.cluster_centers_[:, 0], kmeans.cluster_centers_[:,1], s = 50, c = 'black', label = 'Centroide', marker = '*', alpha=0.5) # for plotting the calculated centroids
    # plt.legend()

    plt.title('Cluster_'+str(number_of_class))
    plt.legend(loc='center left', bbox_to_anchor=(1, 0.5))
    plt.savefig('figure/fig_cluster_'+str(number_of_class)+'.png', bbox_inches="tight") #saving the figures into a folder named "figure".
    plt.clf()
    plt.show()
    return y_kmeans



def main():
    filename ='Datasets.xls' # path to the dataset
    x = read_data(filename)
    #print(x)
    print(x.shape)
   
    for i in range(3,12,2): # running kmean algorithm using different number of clusters.
        model = run_KMean(i,x)
        print('Figure '+str(i)+' Saved...')
    BreakPoint = 999
    


if __name__== "__main__":
    main()

        
        



    
