# ExcelArt

With this tool it is possible to convert your images to awesome ExcelArt!<br>
Due to limitations of Excel the image cannot be converted 1:1. However the number of colours in the image must be reduced by clustering.

![example for conversion](https://github.com/fdahle/ExcelArt/blob/main/examples/example.PNG?raw=true)

(On the left you can see the original image, on the right you can see the image as an excel sheet. Not that the right image is created by colouring the single excel cells, not just by adding the picture)


## How to use it
All necessary code can be found in converter.py. The simplest way of using the program is to call the function `convert_to_excelArt()` with just the imagePath as a variable:

`convert_to_excelArt(imgPath)`

In this case a new Excel-File is created in the current folder with the same name as the image file.

it is possible to use some optional parameters:

* `savePath` - the path and the name of the output excelFile - default: "" (same Folder)
* `overwrite`- if this is set to true it is possible to overwrite exiting excelFiles - default: False
* `pixel_size` - array that specifies the size of the cells in the Excel (\[width, height \]) - default: \[2, 2 \]
* `gray` - if the image should be converted to a grayscale version - default: False
* `scale` - if the image should be rescaled. Must be a float value between 0.001 and 2 - default: 1
* `n_clusters` - the number of clusters in which the image is clustered (=number of different colours in the picture) - default: 128
* `random_state` - random state for clustering in order for reproducibility - default: 0

## How does it work
1. The image is load via matplotlib
2. In order to cluster the image the colorscale is converted to the LAB colorspace<br>
   (necessary for clustering, for more informations see [here](https://www.pyimagesearch.com/2014/07/07/color-quantization-opencv-using-k-means-clustering/)
3. Convert to grayscale if necessary
4. Cluster the image use kmeans
5. Create an Excel-File
6. Iterate through the pixels of the Image and create correspondingly the cells in Excel

## Remarks
* depending on the image size especially the clustering can take quite a while. For the clustering it is currently not possible to display a progress bar, so just be patient. If there are no error messages it is still running
* Excel can support only a limited number of different background-colours at the same time. If the number of clusters is set too high, the Excel-File will crash. In this case just try again with a lower number of clusters.
* Already with 128 different colours you can have a very realistic looking picture. Only for large areas with gradual changing (like a blue sky) you will see patterns.
