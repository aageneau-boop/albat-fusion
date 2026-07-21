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

        candidate = os.path.join(base, *parts)

        # Depuis PyInstaller 6.0, le contenu embarqué (dont
        # assets/, profils/ et "dossiers pour rapport/" listés dans
        # le .spec) est placé par défaut dans un sous-dossier
        # "_internal/" plutôt que directement à côté de
        # l'exécutable — les versions antérieures le mettaient à
        # plat. Plutôt que de dépendre d'un réglage spécifique à
        # une version de PyInstaller (fragile, sujet à changer),
        # on vérifie les deux emplacements possibles : ça fonctionne
        # quelle que soit la version utilisée pour la compilation.
        if not os.path.exists(candidate):
            internal_candidate = os.path.join(base, "_internal", *parts)
            if os.path.exists(internal_candidate):
                return internal_candidate

        return candidate

    base = os.path.dirname(
        os.path.dirname(os.path.abspath(__file__))
    )

    return os.path.join(base, *parts)
