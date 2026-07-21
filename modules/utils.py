import os
import sys


def resource_path(*parts):

    if getattr(sys, "frozen", False):
        # Toujours le dossier réel où se trouve l'exécutable, JAMAIS
        # sys._MEIPASS : avec un packaging --onefile, _MEIPASS
        # pointe vers un dossier temporaire recréé à chaque
        # lancement (et supprimé à la fermeture), ce qui casserait
        # les dossiers personnalisables ("dossiers pour rapport/",
        # "profils/") que l'utilisateur doit pouvoir éditer et
        # retrouver d'une session à l'autre. Nécessite un build
        # PyInstaller en mode --onedir (voir Albat Fusion.spec).
        # Identique à resource_path() dans Albat_Fusion_Mac.py —
        # les deux DOIVENT rester alignés.
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(
            os.path.dirname(os.path.abspath(__file__))
        )

    return os.path.join(base, *parts)