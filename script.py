import argparse
import sys
from threading import Thread, Event
from time import sleep
from traceback import print_exc

from compta import Compta
from devis import Devis
from doli_exceptions import NumCompteNonRempli, DesignationNonRemplie, LienNonReconnu
from note_de_frais import NoteDeFrais


#  process determines from link which task to do (NDF, Devis, Compta...)
def process(url):
    if "expensereport" in url:
        NoteDeFrais(url, username, password, start, end)
    elif "accountancy" in url:
        Compta(url, username, password, start, end)
    elif "propal"in url:
        Devis(url, username, password, start, end)
    else:
        raise LienNonReconnu

def waiting():
    iteration = 0
    animation = [
        '[|     ]',
        '[||    ]',
        '[|||   ]',
        '[||||  ]',
        '[||||| ]',
        '[||||||]',
        '[ |||||]',
        '[  ||||]',
        '[   |||]',
        '[    ||]',
        '[     |]',
        '[      ]'
    ]

    while not thread.stop_event._flag:  # not thread.stop_event
        print("En cours", animation[iteration % len(animation)], end='\r')
        iteration += 1
        sleep(0.1)



# arguments in cmd/powershell
sys.argv = [str(i).lower() for i in sys.argv]
parser = argparse.ArgumentParser(prog='script.py',
                                 usage='python script.py [options] [attr]',
                                 description="Documentation de DoliPy : https://github.com/KataKatis/DoliPy. Ci-dessous, la liste des options de DoliPy",
                                 epilog="N'hésitez pas à faire part de vos bugs / améliorations.")
parser.add_argument('--url', '--site', help='Lien vers Dolibarr', type=str, required=True)
parser.add_argument('--username', help="Nom d'utilisateur",  type=str, required=True)
parser.add_argument('--password', help='Mot de passe', type=str, required=True)
parser.add_argument('--start', help='Ligne de départ dans votre document Excel', type=int, required=True)
parser.add_argument('--end', help='Ligne de fin dans votre document Excel', type=int, required=True)

args = parser.parse_args()

print(sys.argv)

try:
    # wainting animation
    thread = Thread(target=waiting)
    thread.stop_event = Event()  # thread.stop_event._flag = False
    thread.start()

    url = args.url
    username = args.username
    password = args.password
    start = args.start-1
    end = args.end

    # detect process
    process(url)

    thread.stop_event.set()  # thread.stop_event._flag = True

    print("Opération effectuée avec succès !")

except LienNonReconnu:
    thread.stop_event.set()  # thread.stop_event._flag = True
    print("Le lien que vous avez entrez est incorrect.")

except NumCompteNonRempli:
    thread.stop_event.set()  # thread.stop_event._flag = True
    print("Les numéros de comptes sont obligatoires. Vérifiez votre fichier template, dans la feuille 'Compta'.")

except DesignationNonRemplie:
    thread.stop_event.set()  # thread.stop_event._flag = True
    print("La colonne désignation doit être remplie obligatoirement. Vérifiez votre fichier template, dans la feuille 'Devis'.")

except:
    thread.stop_event.set()  # thread.stop_event._flag = True
    print_exc()  # print error even if except statement for debug (remove it for delivering DoliPy)
    print("Une erreur est survenue lors de l'exécution du programme. Réessayer tout de même de l'exécuter, parfois ces erreurs sont dues à un trop long chargement de la page et peuvent se rétablir juste après.")
