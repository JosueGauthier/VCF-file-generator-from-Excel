# -*- coding: utf-8 -*-

# vcard specifications from https://tools.ietf.org/id/draft-ietf-vcarddav-vcardrev-02.html

import pandas as pd

print("\n\n--------------Zapromss de la Boquette d'Angers--------------\n\n")
print("il gnasse un fichier excel (.xlsx) en vcard (.vcf)")
print("(en gros, il t'usine le Zprom'ss numérique à partir du excel remplit par tes chticop'ss)")
print("les champs disponibles sont les suivants : Bucque, Fam'ss, Prénom, Nom, Prom'ss, Blairal(Tel), Mail, Date de naissance")
print("à noter que les champs ''mail'' et ''date de naissance'' ne sont pas nécessaires et peuvent être Zocqués\n\n\n")

# open the patern file and load it
patFile = open("patern.txt", "r")
patern = patFile.read()
patFile.close()



filepath = input("chemin absolu du fichier excel: (sans les guillemets) ")


# constantes pour l'excel (numéro des colonnes des données)
CONSTANTES = {"Bucque":None, "Fam's":None, "Prenom":None, "Nom":None, "Prom's":None, "Mail":None, "Tel":None, "BirthDay":None, "Bouls":None}
print("Quelle prom'ss ? (oüblie pas le tabagn'ss sinon il ne sera pas marqué)")

CONSTANTES["Prom's"] = input("(ex: KIN 220) : ")
print("entrer la colonne correspondante de l'excel (0 si il n'y en a pas)")

for i in CONSTANTES:
    if i != "Prom's":
        CONSTANTES[i] = int(input("colonne {}:".format(i)))

# open the vcard file to write
vcf = open("Zproms {}.vcf".format(CONSTANTES["Prom's"]), "w")


# open the xl file with all the data

xl = pd.read_excel(filepath)

# get the values in a list format
list = xl.values



compteur = 0
# read it and write the corresponding contact to vcard
for line in list:
    print(compteur)
    newContact = patern
    # oula que c'est pas beau
    print(line[CONSTANTES["Bucque"]-1])
    newContact=newContact.replace("bucque", str(line[CONSTANTES["Bucque"]-1]))
    newContact=newContact.replace("fams", str(int(line[CONSTANTES["Fam's"]-1])))
    newContact=newContact.replace("prenom", str(line[CONSTANTES["Prenom"]-1]))
    newContact=newContact.replace("nom", str(line[CONSTANTES["Nom"]-1]))
    newContact=newContact.replace("proms", str(CONSTANTES["Prom's"]))
    newContact=newContact.replace("bouls", str(line[CONSTANTES["Bouls"]-1]))
    # Mail
    if CONSTANTES["Mail"] == 0 or str(line[CONSTANTES["Mail"]-1]) == "nan":
        newContact=newContact.replace("EMAIL;TYPE=internet:mail\n", "")
    else:
        newContact=newContact.replace("mail", str(line[CONSTANTES["Mail"]-1]))

    # Telephone
    if str(line[CONSTANTES["Tel"]-1]) == "nan":
        newContact=newContact.replace("TEL;TYPE=CELL:tel\n", "")
    else:
        telfloat = str(line[CONSTANTES["Tel"]-1])[:] # 6 xx xx xx xx.0 -> 6 xx xx xx xx
        tel = "0"+telfloat
        newContact=newContact.replace("tel", tel)

    # Birth Day traitment:
    if CONSTANTES["BirthDay"] == 0 or str(line[CONSTANTES["BirthDay"]-1]) == "NaT":
        newContact=newContact.replace("BDAY;VALUE=text:bday\n", "")

    else:
        # 2021-05-17 00:00:00   INPUT
        # 20000510 (aaaammdd)   OUTPUT
        bday = str(line[CONSTANTES["BirthDay"]-1]).split(" ")[0]
        bday = bday.replace("-","")
        newContact=newContact.replace("bday", bday)   

    # Boul's works



    if CONSTANTES["Bouls"] == 0 or str(line[CONSTANTES["Bouls"]-1]) == "nan":
        newContact=newContact.replace("bouls1", "bouls1")
    else:

        def FindMaxLength(lst):
            maxList = max(lst, key = len)
            maxLength = max(map(len, lst))
            
            return maxList, maxLength

        list_bouls = []

        sizel = FindMaxLength(list)[1]


        for i in range(0,sizel-7) :

            list_bouls.append(str(line[CONSTANTES["Bouls"]-1+i]))

        cleanedListBouls = [x for x in list_bouls if str(x) != 'nan']


        strbouls = ' | '.join([str(elem) for elem in cleanedListBouls])

                
        newContact=newContact.replace("bouls", strbouls)


    vcf.write(newContact+"\n")
    compteur+=1


vcf.close()

print("le programme a enregistré {} contacts".format(compteur))
