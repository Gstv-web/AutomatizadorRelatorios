import whatsappAtmosfera
import whatsappBewake
import whatsappIknow
import whatsappLigdata
import whatsappAlcance
import os
from time import sleep






while True:
    escolha = int(input("Escolha o tipo de plataforma: \n 1 - Atmosfera \n 2 - Bewake \n 3 - Iknow \n 4 - Ligdata \n 5 - Alcance \n 6 - Sair \n Escolha sua opção: "))

    if escolha == 1:
        whatsappAtmosfera.whatsappAtmosfera()
    elif escolha == 2:
        whatsappBewake.whatsappBewake()
    elif escolha == 3:
        whatsappIknow.whatsappIknow()
    elif escolha == 4:
        whatsappLigdata.whatsappLigdata()
    elif escolha == 5:
        whatsappAlcance.whatsappAlcance()
    elif escolha == 6:
        print("Encerrando...")
        exit()
    else:
        print("Digite uma opção válida! \n ")
        sleep(1.5)