from os import listdir
import os
import openpyxl
from operator import itemgetter

from traitement_excel import (
    folderChoice,
    folderconvert,
    filetest,
    checkfile,
    takeclientlist,
    getcustomersdata
)


def main():
    currentfolderwin = str(folderChoice())
    currentfolder = str(folderconvert(currentfolderwin))

    fileInFolderList = listdir(currentfolder)

    filelist = filetest(fileInFolderList)
    print("les fichiers retenu sont " + str(filelist))

    os.chdir(currentfolderwin)

    checkfile(filelist)

    allCustomersInFolderList = takeclientlist(filelist)

    customersdatalist = getcustomersdata(allCustomersInFolderList, filelist)

    # Nettoyage des valeurs None
    for check in customersdatalist:
        for x in check:
            for i in range(len(x)):
                if x[i] is None:
                    x[i] = 0

    # Tri par date
    for toto in customersdatalist:
        toto.sort(key=itemgetter(1))

    # Création du dossier factureclients
    output_folder = os.path.join(currentfolder, "factureclients")
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    os.chdir(output_folder)

    # Génération des factures
    for data in customersdatalist:
        wbOutput = openpyxl.Workbook()
        sheet = wbOutput.active

        sheet["A1"] = "Facture"
        sheet["B6"] = "DATE"
        sheet["C6"] = "Maxi"
        sheet["D6"] = "Geant"
        sheet["E6"] = "Miche"
        sheet["F6"] = "Galette"
        sheet["G6"] = "Somun"
        sheet["H6"] = "Marguerite"
        sheet["I6"] = "Pide"

        linetable = 7

        for inDATA in data:
            filenameOutput = str(inDATA[0])
            date = inDATA[1].strftime("%d/%m/%Y")

            sheet["A2"] = filenameOutput
            sheet["B" + str(linetable)] = date
            sheet["C" + str(linetable)] = int(inDATA[2])
            sheet["D" + str(linetable)] = int(inDATA[3])
            sheet["E" + str(linetable)] = int(inDATA[4])
            sheet["F" + str(linetable)] = int(inDATA[5])
            sheet["G" + str(linetable)] = int(inDATA[6])
            sheet["H" + str(linetable)] = int(inDATA[7])
            sheet["I" + str(linetable)] = int(inDATA[8])

            linetable += 1

        wbOutput.save(filenameOutput + ".xlsx")


if __name__ == "__main__":
    main()
