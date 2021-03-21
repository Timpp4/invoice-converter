# TODO tuotekoodi ja packsize erottelu, jos ei ole eroteltu välilyönnillä
import pdfplumber
import pandas as pd
import openpyxl
import sys
import glob
import time

class errorClass:
    invoiceName = ""
    invoiceError = ""
    invoiceSuggestion = ""
    


def isInt(s):
    try:
        int(s)
        return True
    except ValueError:
        return False

def readPDF(fileName, dataFrames):
    # Tyhjien datasettien tunnistamiseksi
    dataSets = []

    # Alustus
    clientsLocationName = ""
    invoiceDate = ""
    orderID = ""
    orderDate = ""
    productCode = ""
    product = ""
    groupCode = ""
    packSize = ""
    quantityPurchased = ""
    quantityPurchasedUnit = ""
    netUnitPriceAmount = ""
    netUnitCost = ""
    totalNetCost = ""
    invoiceNumberBoolean = False
    invoiceDateBoolean = False
    tuoteRiviBoolean = False
    productFoundBoolean = False
    riviKoodiBoolean = False
    dataCouldNotBeRead = False

    # Virhetilanteiden seurantaa ja kokeilut

    with pdfplumber.open(fileName) as pdf:

        # PDF erillisiksi sivuiksi
        for page in pdf.pages:
            text = page.extract_text()

            # Sivut riveiksi
            for line in text.split('\n'):

                # Rivi splitataan listaan (whitespace)
                tmp_list = line.split()

                # Käy rivi läpi
                for i in range(len(tmp_list)):

                    # Laskun numero ja päiväys
                    if (invoiceNumberBoolean == False or invoiceDateBoolean == False):
                        if (tmp_list[i] == "nr"):
                            invoiceNumberBoolean = True
                        elif (tmp_list[i] == "Datum"):
                            invoiceDateBoolean = True
                            invoiceDate = tmp_list[i+1]

                    # Tilausnumero, tilauspäivämäärä ja toimitusosoite
                    if (tmp_list[i] == "Tilaus"):
                        if (tmp_list[i+2] == "Order"):
                            tmp_orderNumberDate = tmp_list[i+3]
                            tmp_orderNumberDate = tmp_orderNumberDate.split("(")

                            clientsLocationName = tmp_list[-1]
                            orderID = tmp_orderNumberDate[0]
                            orderDate = tmp_orderNumberDate[1][:-1]

                            print("**********\n" + "Clients Location Name: " + clientsLocationName)
                            print("Tilauspäivämäärä: " + tmp_orderNumberDate[1][:-1])
                            print("Tilausnumero: " + tmp_orderNumberDate[0]  + "\n")

                    # Tuoterivi käyttäen Group-koodia esim. 8205.59.80
                    if (len(tmp_list[i]) == 10 and tmp_list[i][4] == "."):
                        tuoteRiviBoolean = True
                        for data in tmp_list:
                            # Älä lisää Group-koodia tuotenimen loppuun
                            if (data == tmp_list[len(tmp_list)-1]):
                                break
                            # Jos tuote löytyi, koodi jatkaa tästä
                            elif (productFoundBoolean == True):
                                product = product + " " + data
                            # Ensimmäinen alkio, joka ei ole tuotekoodia
                            elif (isInt(data) == False):
                                # Jos tuotekoodissa on kirjaimia, niin pitää olla tarkempi
                                if (data == tmp_list[1]): # Tuote ei voi olla 1. alkio riviltä, on aina koodi
                                    productCode = productCode + data
                                else:
                                    productFoundBoolean = True
                                    if (product == ""):
                                        product = product + data
                                    else:
                                        product = product + " " + data
                            else:
                                if (data == tmp_list[0] and riviKoodiBoolean == False):
                                    riviKoodiBoolean = True # Rivikoodi ja packsize voivat olla samat
                                    continue
                                elif (productCode == ""):
                                    productCode = productCode + data
                                else:
                                    productCode = productCode + " " + data
                        packSize = productCode.split()[-1]
                        groupCode = tmp_list[len(tmp_list)-1]
                        print("Tuote: " + product)
                        print("Group-koodi: " + tmp_list[len(tmp_list)-1])
                        print("Tuotekoodi: " + productCode)
                        print("PackSize: " + packSize)

                        # Alustus seuraavaa tuotetta varten
                        productFoundBoolean = False
                        riviKoodiBoolean = False
                        continue # Jatkaa seuraavalle riville alempaan if-lauseeseen
                    


                    # quantity, gross unit cost, net unit cost, total net cost
                    if (tuoteRiviBoolean == True):
                        tuoteRiviBoolean = False
                        try:
                            print(tmp_list)
                            quantityPurchased = tmp_list[0]
                            quantityPurchasedUnit = tmp_list[1]
                            netUnitCost = tmp_list[2]
                            netUnitPriceAmount = tmp_list[4]
                            totalNetCost = round(int(tmp_list[0]) / int(tmp_list[4]) * float(tmp_list[2].replace(",", ".")), 2)
                        except Exception as e:
                            print(str(e))
                            print("continuing....")
                            print("alustetaan...")
                            virhe = errorClass()
                            virhe.invoiceName = str(fileName)
                            virhe.invoiceError = str(e)
                            virhe.invoiceSuggestion = str("Tarkista tämä tuote ja lisää tarvittaessa:\n\t" + product + "\n\t" + productCode)
                            errorList.append(virhe)



                            
                            productCode = ""
                            product = ""
                            continue

                        print("Quantity Purchased: " + quantityPurchased)
                        print("Quantity Purchased Unit: " + quantityPurchasedUnit)
                        print("Net Unit Price Amount: " + netUnitPriceAmount)          
                        print("Net Unit Cost: " + netUnitCost)
                        print("Total Net Cost: " + str(totalNetCost).replace(".", ",") + "\n")

                        # Uusi dataframe malli
                        df = pd.DataFrame(
                            [[str(fileName), str(clientsLocationName), str(invoiceDate), str(orderID), str(orderDate), str(productCode), str(product), str(groupCode),
                            str(packSize), str(quantityPurchased), str(quantityPurchasedUnit), str(netUnitPriceAmount), str(netUnitCost), str(totalNetCost)]],
                            columns=['InvoiceName', 'ClientsLocationName', 'InvoiceDate', 'OrderNumber', 'OrderDate', 'ProductCode', 'Description',
                            'GroupCode', 'PackSize', 'QuantityPurchased', 'QuantityPurchasedUnit', 'NetUnitPriceAmount', 'NetUnitCost', 'Total@NetCost']
                        )
                        # Lisää malli listaan
                        dataFrames.append(df)
                        dataSets.append(df) # Virheiden hallintaa varten

                        # Alustus seuraavaa tuotetta varten
                        productCode = ""
                        product = ""
                        
    # Lasku käyty läpi, palauta tiedot main()
    if not dataSets: # Jos laskulta ei saatu dataa, annetaan virheilmoitus
        dataCouldNotBeRead = True
        return dataFrames, dataCouldNotBeRead
    return dataFrames, dataCouldNotBeRead # Palauta data ja tieto onnistumisesta


def main():
    # Laskujen data tähän listaan
    toExcelList = []
    # Virheiden kaappaus ja tunnistus
    errorCheck = False
    # Nykyinen työkansio, sis. kaikki .pdf -tiedostot
    path = '*.pdf'
    files = glob.glob(path)
    for name in files:
        try:
            toExcelList, errorCheck = readPDF(name, toExcelList)
            if (errorCheck == True): # Jos laskulta ei saatu dataa
                virhe = errorClass()
                virhe.invoiceName = str(name)
                virhe.invoiceError = "Dataa ei voitu lukea"
                virhe.invoiceSuggestion = "Käy laskut läpi käsin"
                errorList.append(virhe)

        except Exception as e: # Jokin muu virhe laskua luettaessa (esim. väärän firman lasku)
            virhe = errorClass()
            virhe.invoiceName = str(name)
            virhe.invoiceError = str(e)
            virhe.invoiceSuggestion = "Laskua voitiin lukea, mutta sitä ei voitu jäsentää"
            errorList.append(virhe)


    # Yhdistetään listassa oleva data
    try:
        result = pd.concat(toExcelList)
    except Exception as e:
        print("Odottamaton virhe dataa yhdistäessä: " + str(e))
        print("Virheitä luettaessa laskuja: " + str(len(errorList)))
        for x in errorList:
            print("\t" + x.invoiceName)
            print("\t" + x.invoiceError)
            print("\t" + x.invoiceSuggestion + "\n")
        return
        

    # Kirjoitetaan data excelliin
    try:
        result.to_excel("out.xlsx", index=False)
        print("out.xlsx kirjoitettiin onnistuneesti")
        print("Virheitä luettaessa laskuja: " + str(len(errorList)))
        for x in errorList:
            print("\t" + x.invoiceName)
            print("\t" + x.invoiceError)
            print("\t" + x.invoiceSuggestion + "\n")
    
    except Exception as e:
        print("Odottamaton virhe kirjoittaessa exceliin: " + str(e))
        print("Virheitä luettaessa laskuja: " + str(len(errorList)))
        for x in errorList:
            print("\t" + x.invoiceName)
            print("\t" + x.invoiceError)
            print("\t" + x.invoiceSuggestion + "\n")


start_time = time.time()
errorList = []

main()
print("\n************* EOF *************")
print("**** Run time %s seconds ****" % round((time.time() - start_time), 2))
print("*******************************\n")
