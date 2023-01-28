#NOTE: FILES THAT ARE READ AND PRODUCED CAN NOT BE OPEN WHILE PROGRAM IS RUN
import pandas as pd


#Main function
#First: data is read from specified file (Must be in same folder as the python script)
#Second: images are matched together and put in an image list (nucleoli of individual are grouped together)
#Third: Cells are found while data is still grouped by image
#Fourth: Supplementary functions are run and the data is written to excel

def main():
    nucleoliData = read_data("3&4MaskedNucleoli")
    images = matchImages(nucleoliData)
    cells = matchCells(images)
    toExcel(cells)


#Data is read by specified filename, and AreaShape_Area, Number_Object_Number, and Parent_Masked Nuclei
#columns are isolated and returned to user as a primitive list
def read_data(filename):
    nucleoliDataFrame = pd.read_excel(filename+".xlsx")
    nucleoliData = pd.DataFrame(data = {"AreaShape_Area":nucleoliDataFrame['AreaShape_Area'].values.tolist(),
                                        "Number_Object_Number":nucleoliDataFrame["Number_Object_Number"].values.tolist(),
                                        "Parent_MaskedNuclei":nucleoliDataFrame["Parent_MaskedNuclei"].values.tolist()})

    nucleoliData = nucleoliData.values.tolist()

    return nucleoliData

#Divides nucleoli data based on image taken
def matchImages(data):
    #Initialize image lists
    imageBoundary = 0
    images = []
    image = []

    #Loops through data using comparison to determine when to split an image
    for nucleoli in data:
        if(nucleoli[1]>imageBoundary):
            image.append(nucleoli)
            imageBoundary = nucleoli[1]
        
        else:
            images.append(image)
            image = []
            imageBoundary = 0
    
    return images

#Matches Cells based on the imageList, creating a comprehensive acccessible list of dictionary of lists
def matchCells(images):

    totalCells = []
    for image in images:
        imageCells = {}
        for nucleoli in image:
            if nucleoli[2] in imageCells.keys():
                imageCells[nucleoli[2]].append(nucleoli)
            else:
                imageCells[nucleoli[2]] = [nucleoli]

        totalCells.append(imageCells)

    return totalCells

def findNumberOfImages(CellArray):
    return len(CellArray)

def findNumberOfCellsPerImage(CellArray):
    cellCounts = []
    for image in CellArray:
        cells = 0
        for key in image.keys():
            cells += 1
        cellCounts.append(cells)
    
    return cellCounts

def findNumberOfCells(CellArray):
    return sum(findNumberOfCellsPerImage(CellArray))

def findNucleoliPerCell(CellArray):
    cells = []
    for image in CellArray:
        for key in image.keys():
            cells.append(len(image[key]))

    return cells

def findNucleoliAreaPerCell(CellArray):
    cellArea = []
    for image in CellArray:
        for key in image.keys():
            area = 0
            for i in range(len(image[key])):
                area += image[key][i][0]
            cellArea.append(area)

    return cellArea
    
def populateEmptyArray(num):
    numList = []
    for i in range(1,num+1):
        numList.append(i)
    
    return numList

def findImageNum(cellArray):
    imageNums = []
    count = 0
    for image in cellArray:
        count += 1
        for key in image.keys():
            imageNums.append(count)

    return imageNums

def toExcel(cellArray):
    numList = populateEmptyArray(findNumberOfCells(cellArray))

    d = {"Cell":numList,
         "Nucleoli Area":findNucleoliAreaPerCell(cellArray),
         "# of Nucleoli":findNucleoliPerCell(cellArray),
         "Image Taken":findImageNum(cellArray)}

    cellData = pd.DataFrame(data=d)
    cellData.to_excel("MaskedNucleoliAnalysis.xlsx", index = False)

main()