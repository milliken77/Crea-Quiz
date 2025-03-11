import random
import numpy
import sys

#Sintassi : python3 crea_quiz.py filedomande versioni risultatidomande correttore numero
#
#			filedomande : elenco delle domande formattate con la seguente sintassi
#					#Domanda 1
#					Risposta 1
#					*Risposta 2
#					Risposta 3
#					#Domanra 2
#					Risposta 1
#					Risposta 2
#					*Risposta 3
#					...
#					La domanda deve iniziare con un # e la risposta giusta nell'elenco deve cominciare con un *
#
#			versioni : il numero di rimescolamenti diversi di tutte le domande
#			risultatidomande : il file txt con le versioni volute delle domande rimescolate
#			correttore : il file csv dei correttori di ogni versione creata
#                       numero : il numero di domande all'interno di ogni quiz. Se omesso è pari al totale delle domande presenti


versioni=int(sys.argv[2])

def mescola_coppia(a, b):
  joined_lists = list(zip(a, b))
  random.shuffle(joined_lists) # Shuffle "joined_lists" in place
  a,b = zip(*joined_lists) # Undo joining
  return a,b

completo=[]


with open(sys.argv[1]) as file:
  completo = [i.strip() for i in file]

domande=[]
risposte=[]
n_domande=0
for i in completo:
  if i!="":
    if i[0]=="#":
      domande.append(i)
      if n_domande!=0:
        risposte.append(lista)
      lista=[]
      n_domande+=1
    if i[0]!="#":
      lista.append(i)
risposte.append(lista)


file1=open(sys.argv[3], 'w')
file2=open(sys.argv[4], 'w')

numero=int(sys.argv[5])

#Crea le varie versioni
for i in range(0,versioni):
  correttore=[]
  domande,risposte=mescola_coppia(domande,risposte)
  tot=0
  print("Compito n° ",i+1)
  file1.write("Compito n° "+str(i+1)+"\n")
  print("   ")
  file1.write("\n")
  for j in domande[:numero]:
    print(tot+1," : ",j[1:])
    file1.write(str(tot+1)+" - "+j[1:]+"\n")
    #file1.write("\n")
    random.shuffle(risposte[tot])
    risp=0
    for k in risposte[tot]:
      lettera=chr(ord("A")+risp)
      if k[0]=="*":
        k=k[1:]
        correttore.append(lettera)
      print("	",lettera,"  ",k)
      file1.write("    "+lettera+" : "+k+"\n")
      #file1.write("\n")
      risp=risp+1
    tot=tot+1
    #file1.write("\n")
  print(str(i+1)+","+str(correttore)[1:-1].replace("'", ""))
  file2.write(str(i+1)+","+str(correttore)[1:-1].replace("'", "")+"\n")
  print("   ")
  file1.write("\n")
  file1.write("\n")
  file1.write(":::::\n")


file1.close()
file2.close()
