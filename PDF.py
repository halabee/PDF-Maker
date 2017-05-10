import csv
from PyPDF2 import PdfFileReader, PdfFileMerger
from os import listdir, chdir

pathFolder = r"/Users/mhalabi/Desktop/Python/" #update...duh
pathExcel = r"/Users/mhalabi/Desktop/Python/tar.csv" #update...duh
pathAP = r"/Users/mhalabi/Desktop/Finance/AP/AP FY17" #update...duh

#list of all the invoices in the AP Folder
filelist = listdir(pathAP)

#where we place selected files for merging - these are items which
#are in both my targeting file and my AP directory
mergeTargets = []

#list of targets from excel file
targets = []

#open file, 
with open(pathExcel, newline ='') as filetoread:
    openfile = csv.reader(filetoread)
    for p in list(openfile):
        targets.append(p)

flattenedTarList = [item for sublist in targets for item in sublist]

missList = []

for tar in flattenedTarList:
    for filename in filelist:
        if str(tar).lower() in str(filename).lower() and "approval" not in str(filename).lower():
            mergeTargets.append(filename)
            break
            #break so you only find one in case it duplicates - if
            #you WANT the dups - A) Fix naming conventions, and
            #B)remove break :)


        
#This goes through because the above loop creating the merge targets list
#goes through and will find a fail for every "tar" item until it succeeds
#so every item will end up "failing" a few times unless it's the first
#item in the filelist.

#this block just creates the class so we can write PDF and
#sets the directory so the PDF is created in our directory folder
#that we want it written to (in this case the folder the tar file is in
outfile = PdfFileMerger()
chdir(pathFolder)

mergeTargets.sort(reverse=True)

print(mergeTargets)
for filename in mergeTargets:
    p = pathAP + "/" + filename
    outfile.append(PdfFileReader(p, 'rb'))
    print("Appending file: " + str(filename))

print("Printing to file: Merged PDF.pdf")
outfile.write("Merged PDF.pdf")



