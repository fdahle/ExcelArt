import os
import warnings
import sys

import cv2
import xlsxwriter
import numpy as np

from datetime import datetime
from matplotlib.image import imread
from sklearn.cluster import MiniBatchKMeans

def convert_to_excelArt(imgPath, savePath="", overwrite=False, pixel_size = [2, 2], gray=False, scale=1, n_clusters = 128, random_state=0):

    #check if imgPath is correct
    if (os.path.exists(imgPath) == False):
        raise Exception("No file could be found at the specified location.")

    #if no savePath is specified the file will be saved in the same folder with the
    #name of the image file
    if savePath == "":

        #get name of file
        fileName = os.path.splitext(imgPath)[0]

        #get current date and time with underscore
        now = datetime.now()
        timestamp = now.strftime("%Y_%m_%d_%H_%M_%S")

        #create savePath
        savePath = fileName + "_" + timestamp + ".xlsx"

    else:
        folder = os.path.dirname(savePath)
        if (folder != "" and os.path.exists(folder) == False):
            raise Exception("The saving folder does not exist.")
        if (overwrite == False and os.path.exists(savePath) == True):
                raise Exception("The file already exists.")
        if (os.path.splitext(savePath)[1] != ".xlsx"):
            raise Exception("The file type must be .xlsx")

    if scale < 0.001:
        raise Exception("Minimum scale allowed is 0.001")
    if scale > 2:
        raise Exception("Maximum scale allowed be 2")

    if n_clusters < 0:
        raise Exception("At least 0 (for no clustering) or 1 cluster (for clustering) must be entered.")
    elif n_clusters > 1024:
        warnings.warn("With this number of different colours the excel-file may be corrupted.")

    if (len(pixel_size) != 2):
        raise Exception("pixel_size must be an array with two integers.")
    if (pixel_size[0] < 1):
        pixel_size[0] = 1
    if (pixel_size[1] < 1):
        pixel_size[1] = 1

    #specify where the image is
    imgPath = imgPath

    #read image as numpy array
    image = imread(imgPath)

    #down- or upscale
    new_width = int(image.shape[1] * scale)
    new_height = int(image.shape[0] * scale)
    dim = (new_width, new_height)
    image = cv2.resize(image, dim)

    #make to gray and set back to rgb for clustering (but will still have the gray values)
    if gray:
        image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        image = cv2.cvtColor(image,cv2.COLOR_GRAY2RGB)

    #cluster image
    def clusterImage(img, n_clusters):

        #based on https://www.pyimagesearch.com/2014/07/07/color-quantization-opencv-using-k-means-clustering/
        height, width = img.shape[:2]

        #another colorspace
        img_lab = cv2.cvtColor(img, cv2.COLOR_BGR2LAB)

        #convert to a feature vector
        feat_vec = img_lab.reshape((img.shape[0] * img.shape[1], 3))

        #apply clustering
        k_means = MiniBatchKMeans(n_clusters=n_clusters, random_state=random_state)
        k_means.fit(feat_vec)

        #replace the values of the img with the cluster centers
        labels = k_means.labels_
        clustered_img = k_means.cluster_centers_.astype("uint8")[labels]

        #reshape and convert to real rgb image
        clustered_img = clustered_img.reshape((height, width, 3))
        clustered_img = cv2.cvtColor(clustered_img, cv2.COLOR_LAB2BGR)

        return clustered_img
    if n_clusters > 0:
        print("Image is clustered..")
        image = clusterImage(image, n_clusters)



    #get the workbook and sheet
    workbook = xlsxwriter.Workbook(savePath)
    worksheet = workbook.add_worksheet()

    #replaces the number with the excel letter
    def colnum_string(n):
        string = ""
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            string = chr(65 + remainder) + string
        return string

    #convert rgb to hex
    def rgb_to_hex(rgb):
        return '#%02x%02x%02x' % rgb

    #saves the already created colour-formats
    hex_dict = {};

    total_num_pixels = image.shape[0] * image.shape[1]
    count_pixels = 0

    print("ExcelArt will be created:")

    #iterate the image
    for row in range(image.shape[0]):
        for col in range(image.shape[1]):

            #increment percentage
            count_pixels = count_pixels + 1
            percentage = int(count_pixels/total_num_pixels * 100)
            sys.stdout.write("\r[%-100s] %d%%" % ('='*percentage, percentage))
            #sys.stdout.write("\r" + str(){0}>".format("="*percentage))
            sys.stdout.flush()

            #get color and convert to hex
            rgb = image[row, col]
            hex = rgb_to_hex((rgb[0], rgb[1], rgb[2]))

            #if colour already existing no need to create a new format
            if hex in hex_dict:
                cell_format = hex_dict[hex]

            #create a new format
            else:

                #set the format for this cell
                cell_format = workbook.add_format()
                cell_format.set_bg_color(hex)
                cell_format.set_pattern(1)
                hex_dict[hex] = cell_format

            #get the excel position
            excel_pos = colnum_string(col+1) + str(row+1)

            #write this cell to the excel
            worksheet.write(excel_pos, '', cell_format)

    #set the size of the columns
    worksheet.set_column(0, new_width - 1, pixel_size[0]/12)
    worksheet.set_default_row(pixel_size[1])
    #close and save the workbook
    workbook.close()


convert_to_excelArt(imgPath, savePath, overwrite=True)
