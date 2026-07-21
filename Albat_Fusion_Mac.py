import sys
import os
import json
import shutil
import subprocess
import datetime as _dt
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QTabWidget,
    QDialog, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton,
    QFileDialog, QMessageBox, QWidget,
    QInputDialog, QTextEdit, QComboBox,
    QSplashScreen
)
from PySide6.QtGui import QPixmap
from PySide6.QtCore import Qt, QTimer

# Les imports des modules (lourds : pandas, matplotlib, python-docx...)
# sont volontairement faits plus tard, dans Suite.__init__, une fois
# l'écran de chargement affiché, plutôt qu'ici au niveau du module.

PROJECT_SUBFOLDERS = [
    "Correlations", "Graph", "Bridage", "Scenar", "Rapport", "Qgis"
]


def ensure_project_subfolders(project_dir):
    """Crée (si besoin) les 5 sous-dossiers standards d'un projet
    Albat, un par onglet."""

    for sub in PROJECT_SUBFOLDERS:
        os.makedirs(os.path.join(project_dir, sub), exist_ok=True)
def resource_path(*parts):
    if getattr(sys, "frozen", False):
        base = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    else:
        base = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base, *parts)


PROFILS_DIR = resource_path("profils")


def list_bureau_profiles():
    """
    Retourne un dict {nom_bureau: chemin_du_.json} pour chaque
    profil de bureau d'étude trouvé dans le dossier global
    'profils' (un sous-dossier par bureau).
    """

    profiles = {}

    if not os.path.isdir(PROFILS_DIR):
        return profiles

    for sub in sorted(os.listdir(PROFILS_DIR)):

        sub_path = os.path.join(PROFILS_DIR, sub)

        if not os.path.isdir(sub_path):
            continue

        for f in os.listdir(sub_path):
            if f.lower().endswith(".json"):
                profiles[sub] = os.path.join(sub_path, f)
                break

    return profiles


class BureauProfileDialog(QDialog):
    """
    Boîte de dialogue pour créer un nouveau profil de bureau
    d'étude (logo, coordonnées), enregistré dans le dossier global
    'profils' pour être réutilisable dans tous les futurs projets.
    """

    def __init__(self):
        super().__init__()

        self.logo_path = None
        self.logo_secondaire_path = None
        self.created_name = None

        self.setWindowTitle("Créer un bureau d'étude")
        self.setMinimumWidth(420)

        from PySide6.QtGui import QIcon
        icon_path = resource_path("assets", "icone_albat.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        self.setStyleSheet("""
            QDialog{ background:#12210f; }
            QLabel{ color:#eef7d0; font-size:12px; font-weight:600; }
            QLineEdit, QTextEdit{
                background:black;
                color:#eef7d0;
                border-radius:10px;
                padding:8px;
                border:1px solid rgba(255,255,255,45);
                font-size:13px;
            }
            QPushButton{
                background:rgba(90,127,71,200);
                color:white;
                border-radius:12px;
                padding:8px 16px;
                font-size:12px;
                font-weight:700;
            }
            QPushButton:hover{
                background:rgba(120,160,95,230);
            }
        """)

        layout = QVBoxLayout(self)
        layout.setSpacing(8)
        layout.setContentsMargins(20, 20, 20, 20)

        layout.addWidget(QLabel("Nom du bureau d'étude :"))
        self.field_nom = QLineEdit()
        self.field_nom.setPlaceholderText("ex : Engoulevent Ecologie")
        layout.addWidget(self.field_nom)

        # --- Logo principal ---
        logo_row = QHBoxLayout()
        logo_row.setSpacing(10)

        self.logo_preview = QLabel("Aucun logo")
        self.logo_preview.setFixedSize(80, 80)
        self.logo_preview.setAlignment(Qt.AlignCenter)
        self.logo_preview.setStyleSheet("""
            background:rgba(0,0,0,0.35);
            border:1px solid rgba(255,255,255,60);
            border-radius:8px;
            color:#cccccc;
            font-size:10px;
        """)

        btn_logo = QPushButton("Choisir un logo")
        btn_logo.clicked.connect(self.pick_logo)

        logo_row.addWidget(self.logo_preview)
        logo_row.addWidget(btn_logo, 1)

        layout.addLayout(logo_row)

        # --- Logo secondaire ---
        logo2_row = QHBoxLayout()
        logo2_row.setSpacing(10)

        self.logo2_preview = QLabel("Aucun logo")
        self.logo2_preview.setFixedSize(80, 80)
        self.logo2_preview.setAlignment(Qt.AlignCenter)
        self.logo2_preview.setStyleSheet("""
            background:rgba(0,0,0,0.35);
            border:1px solid rgba(255,255,255,60);
            border-radius:8px;
            color:#cccccc;
            font-size:10px;
        """)

        btn_logo2 = QPushButton("Choisir un logo secondaire")
        btn_logo2.clicked.connect(self.pick_logo_secondaire)

        logo2_row.addWidget(self.logo2_preview)
        logo2_row.addWidget(btn_logo2, 1)

        layout.addLayout(logo2_row)

        layout.addWidget(QLabel("Adresse :"))
        self.field_adresse = QTextEdit()
        self.field_adresse.setPlaceholderText(
            "ex : Le Champ de la Cure, 58230 Saint Agnan"
        )
        self.field_adresse.setFixedHeight(60)
        layout.addWidget(self.field_adresse)

        layout.addWidget(QLabel("Téléphone :"))
        self.field_tel = QLineEdit()
        layout.addWidget(self.field_tel)

        layout.addWidget(QLabel("Email :"))
        self.field_email = QLineEdit()
        layout.addWidget(self.field_email)

        layout.addWidget(QLabel("Site web :"))
        self.field_web = QLineEdit()
        layout.addWidget(self.field_web)

        btn_row = QHBoxLayout()
        btn_row.addStretch()

        cancel_btn = QPushButton("Annuler")
        cancel_btn.setStyleSheet("""
            QPushButton{ background:rgba(60,60,60,200); }
            QPushButton:hover{ background:rgba(90,90,90,230); }
        """)
        cancel_btn.clicked.connect(self.reject)

        save_btn = QPushButton("Créer le bureau")
        save_btn.clicked.connect(self.save_and_close)

        btn_row.addWidget(cancel_btn)
        btn_row.addWidget(save_btn)

        layout.addLayout(btn_row)

    # ======================================================

    def pick_logo(self):

        path, _ = QFileDialog.getOpenFileName(
            self, "Choisir le logo", "", "Images (*.png *.jpg *.jpeg)"
        )

        if not path:
            return

        pix = QPixmap(path)

        if pix.isNull():
            return

        self.logo_path = path
        self.logo_preview.setPixmap(
            pix.scaled(
                self.logo_preview.size(),
                Qt.KeepAspectRatio,
                Qt.SmoothTransformation
            )
        )

    # ======================================================

    def pick_logo_secondaire(self):

        path, _ = QFileDialog.getOpenFileName(
            self, "Choisir le logo secondaire", "",
            "Images (*.png *.jpg *.jpeg)"
        )

        if not path:
            return

        pix = QPixmap(path)

        if pix.isNull():
            return

        self.logo_secondaire_path = path
        self.logo2_preview.setPixmap(
            pix.scaled(
                self.logo2_preview.size(),
                Qt.KeepAspectRatio,
                Qt.SmoothTransformation
            )
        )

    # ======================================================

    def save_and_close(self):

        nom = self.field_nom.text().strip()

        if not nom:
            QMessageBox.warning(
                self, "Nom manquant",
                "Indiquez le nom du bureau d'étude."
            )
            return

        safe_name = "".join(
            c if c.isalnum() or c in " _-" else "_"
            for c in nom
        ).strip()

        if not safe_name:
            QMessageBox.warning(
                self, "Nom invalide",
                "Ce nom ne peut pas être utilisé."
            )
            return

        bureau_dir = os.path.join(PROFILS_DIR, safe_name)

        try:
            os.makedirs(bureau_dir, exist_ok=True)
        except Exception as e:
            QMessageBox.critical(
                self, "Erreur",
                f"Impossible de créer le dossier du profil :\n{e}"
            )
            return

        logo_filename = None
        logo2_filename = None

        try:
            if self.logo_path and os.path.exists(self.logo_path):
                ext = os.path.splitext(self.logo_path)[1] or ".png"
                logo_filename = f"{safe_name}_logo{ext}"
                shutil.copy(
                    self.logo_path,
                    os.path.join(bureau_dir, logo_filename)
                )

            if (
                self.logo_secondaire_path
                and os.path.exists(self.logo_secondaire_path)
            ):
                ext = (
                    os.path.splitext(self.logo_secondaire_path)[1]
                    or ".png"
                )
                logo2_filename = f"{safe_name}_logo2{ext}"
                shutil.copy(
                    self.logo_secondaire_path,
                    os.path.join(bureau_dir, logo2_filename)
                )

            data = {
                "nom": nom,
                "adresse": self.field_adresse.toPlainText().strip(),
                "telephone": self.field_tel.text().strip(),
                "email": self.field_email.text().strip(),
                "site_web": self.field_web.text().strip(),
                "logo_fichier": logo_filename,
                "logo_secondaire_fichier": logo2_filename,
            }

            with open(
                os.path.join(bureau_dir, f"{safe_name}.json"),
                "w", encoding="utf-8"
            ) as f:
                json.dump(data, f, ensure_ascii=False, indent=2)

        except Exception as e:
            QMessageBox.critical(
                self, "Erreur",
                f"Impossible d'enregistrer le profil :\n{e}"
            )
            return

        self.created_name = nom

        self.accept()


class StartupDialog(QDialog):
    """
    Boîte de dialogue affichée au lancement de l'application :
    permet de créer un nouveau projet (avec choix de l'emplacement,
    ce qui génère automatiquement les dossiers Correlations/Graph/
    Bridage/Scenar/Rapport) ou d'en ouvrir un déjà existant.
    """

    def __init__(self):
        super().__init__()

        self.project_path = None
        self.mode = "create"
        self.bureau_profiles = {}

        self.setWindowTitle("Albat Fusion")
        self.setMinimumWidth(440)

        from PySide6.QtGui import QIcon
        icon_path = resource_path("assets", "icone_albat.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        self.setStyleSheet("""
            QDialog{
                background:#12210f;
            }
            QLabel{
                color:#eef7d0;
                font-size:13px;
                font-weight:600;
            }
            QLineEdit{
                background:black;
                color:#eef7d0;
                border-radius:10px;
                padding:8px;
                border:1px solid rgba(255,255,255,45);
                font-size:13px;
            }
            QPushButton{
                background:rgba(90,127,71,200);
                color:white;
                border-radius:12px;
                padding:10px 20px;
                font-size:13px;
                font-weight:700;
            }
            QPushButton:hover{
                background:rgba(120,160,95,230);
            }
        """)

        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(20, 20, 20, 20)

        title = QLabel("Bienvenue dans Albat Fusion")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size:17px; font-weight:800;")
        layout.addWidget(title)

        # ==========================================
        # CHOIX DU MODE
        # ==========================================

        mode_row = QHBoxLayout()
        mode_row.setSpacing(10)

        self.btn_mode_create = QPushButton("✨ Créer un nouveau projet")
        self.btn_mode_create.clicked.connect(
            lambda: self.set_mode("create")
        )

        self.btn_mode_open = QPushButton("📂 Ouvrir un projet existant")
        self.btn_mode_open.clicked.connect(
            self._on_click_mode_open
        )

        mode_row.addWidget(self.btn_mode_create)
        mode_row.addWidget(self.btn_mode_open)

        layout.addLayout(mode_row)

        # ==========================================
        # BLOC "CRÉER UN PROJET"
        # ==========================================

        self.create_widget = QWidget()
        create_layout = QVBoxLayout(self.create_widget)
        create_layout.setContentsMargins(0, 6, 0, 0)
        create_layout.setSpacing(8)

        create_layout.addWidget(QLabel("Nom du parc éolien :"))
        self.park_input = QLineEdit()
        self.park_input.setPlaceholderText("ex : Vilpion")
        create_layout.addWidget(self.park_input)

        create_layout.addWidget(
            QLabel("Ville de référence (lever / coucher du soleil) :")
        )
        self.city_input = QLineEdit()
        self.city_input.setPlaceholderText("ex : Marle")
        create_layout.addWidget(self.city_input)

        create_layout.addWidget(QLabel("Année :"))
        self.year_input = QLineEdit()
        self.year_input.setPlaceholderText("ex : 2025")
        self.year_input.setText(str(_dt.date.today().year))
        create_layout.addWidget(self.year_input)

        create_layout.addWidget(QLabel("Nom du bureau d'étude :"))

        bureau_row = QHBoxLayout()
        bureau_row.setSpacing(8)

        self.bureau_input = QComboBox()
        self.bureau_input.setEditable(True)
        self.bureau_input.lineEdit().setPlaceholderText(
            "ex : Engoulevent Ecologie"
        )
        self.bureau_input.setStyleSheet("""
            QComboBox{
                background:black;
                color:#eef7d0;
                border-radius:10px;
                padding:6px;
                border:1px solid rgba(255,255,255,45);
                font-size:13px;
            }
            QComboBox QAbstractItemView{
                background:#1a1a1a;
                color:#eef7d0;
                selection-background-color:rgba(90,127,71,220);
                selection-color:white;
                border:1px solid rgba(255,255,255,45);
                outline:0;
            }
        """)

        btn_new_bureau = QPushButton("➕ Créer")
        btn_new_bureau.setFixedWidth(90)
        btn_new_bureau.clicked.connect(self.create_bureau)

        bureau_row.addWidget(self.bureau_input, 1)
        bureau_row.addWidget(btn_new_bureau)

        create_layout.addLayout(bureau_row)

        self.refresh_bureau_list()

        sep = QLabel(
            "Le projet créera automatiquement 5 dossiers "
            "(Correlations, Graph, Bridage, Scenar, Rapport) à "
            "l'emplacement choisi ci-dessous."
        )
        sep.setWordWrap(True)
        sep.setStyleSheet(
            "font-weight:400; font-size:11px; color:#cfe0a0;"
        )
        create_layout.addWidget(sep)

        create_layout.addWidget(QLabel("Nom du projet :"))
        self.project_name_input = QLineEdit()
        self.project_name_input.setPlaceholderText(
            "ex : Albat_Vilpion_2025"
        )
        create_layout.addWidget(self.project_name_input)

        location_row = QHBoxLayout()
        location_row.setSpacing(8)

        btn_location = QPushButton("Choisir l'emplacement")
        btn_location.clicked.connect(self.pick_location)

        self.lbl_location = QLabel("Aucun emplacement choisi")
        self.lbl_location.setWordWrap(True)
        self.lbl_location.setStyleSheet(
            "font-weight:400; font-size:10px; color:#a0c8ff;"
        )

        location_row.addWidget(btn_location)
        location_row.addWidget(self.lbl_location, 1)

        create_layout.addLayout(location_row)

        layout.addWidget(self.create_widget)

        # ==========================================
        # BLOC "OUVRIR UN PROJET"
        # ==========================================

        self.open_widget = QWidget()
        open_layout = QVBoxLayout(self.open_widget)
        open_layout.setContentsMargins(0, 6, 0, 0)
        open_layout.setSpacing(8)

        load_project_btn = QPushButton("Choisir le fichier du projet")
        load_project_btn.setStyleSheet("""
            QPushButton{
                background:rgba(70,100,140,200);
            }
            QPushButton:hover{
                background:rgba(95,130,175,230);
            }
        """)
        load_project_btn.clicked.connect(self.pick_project)
        open_layout.addWidget(load_project_btn)

        self.lbl_project_status = QLabel("Aucun projet sélectionné")
        self.lbl_project_status.setWordWrap(True)
        self.lbl_project_status.setStyleSheet(
            "font-weight:400; font-size:10px; color:#a0c8ff;"
        )
        open_layout.addWidget(self.lbl_project_status)

        layout.addWidget(self.open_widget)

        # ==========================================
        # BOUTONS BAS
        # ==========================================

        btn_row = QHBoxLayout()
        btn_row.addStretch()

        skip_btn = QPushButton("Passer")
        skip_btn.setStyleSheet("""
            QPushButton{
                background:rgba(60,60,60,200);
            }
            QPushButton:hover{
                background:rgba(90,90,90,230);
            }
        """)
        skip_btn.clicked.connect(self.reject)

        self.ok_btn = QPushButton("Valider")
        self.ok_btn.clicked.connect(self.try_accept)

        btn_row.addWidget(skip_btn)
        btn_row.addWidget(self.ok_btn)

        layout.addLayout(btn_row)

        self.chosen_location = None

        self.set_mode("create")

    # ======================================================

    def set_mode(self, mode):

        self.mode = mode

        self.create_widget.setVisible(mode == "create")
        self.open_widget.setVisible(mode == "open")

        active_style = """
            QPushButton{
                background:rgba(130,184,99,230);
                color:white;
                border-radius:12px;
                padding:10px 20px;
                font-size:13px;
                font-weight:700;
            }
        """

        self.btn_mode_create.setStyleSheet(
            active_style if mode == "create" else ""
        )
        self.btn_mode_open.setStyleSheet(
            active_style if mode == "open" else ""
        )

    # ======================================================

    def _on_click_mode_open(self):
        """
        Au clic sur "Ouvrir un projet existant" : bascule l'écran
        en mode "open" ET ouvre directement l'explorateur de
        fichiers pour choisir le .albatproj, au lieu de forcer un
        clic supplémentaire sur "Choisir le fichier du projet".
        Cette dernière reste visible/cliquable pour recommencer si
        l'utilisateur annule ou veut changer de fichier.
        """

        self.set_mode("open")
        self.pick_project()

    # ======================================================

    def refresh_bureau_list(self, select_name=None):

        self.bureau_profiles = list_bureau_profiles()

        current_text = self.bureau_input.currentText()

        self.bureau_input.blockSignals(True)
        self.bureau_input.clear()
        self.bureau_input.addItems(sorted(self.bureau_profiles.keys()))
        self.bureau_input.blockSignals(False)

        if select_name:
            self.bureau_input.setCurrentText(select_name)
        elif current_text:
            self.bureau_input.setCurrentText(current_text)

    # ======================================================

    def create_bureau(self):

        dlg = BureauProfileDialog()

        if dlg.exec() == QDialog.Accepted and dlg.created_name:
            self.refresh_bureau_list(select_name=dlg.created_name)

    # ======================================================

    def pick_location(self):

        path = QFileDialog.getExistingDirectory(
            self, "Choisir le dossier où créer le projet"
        )

        if not path:
            return

        self.chosen_location = path
        self.lbl_location.setText(path)

    # ======================================================

    def pick_project(self):

        manifest_path, _ = QFileDialog.getOpenFileName(
            self,
            "Choisir le fichier du projet Albat",
            "",
            "Projet Albat (*.albatproj);;Ancien format (*.json)"
        )

        if not manifest_path:
            return

        try:
            with open(manifest_path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            QMessageBox.critical(
                self, "Erreur",
                f"Impossible de lire ce projet :\n{e}"
            )
            return

        if "park_name" not in data and "bureau_etude" not in data:
            QMessageBox.warning(
                self, "Fichier invalide",
                "Ce fichier ne semble pas être un projet Albat valide."
            )
            return

        # Le dossier contenant ce fichier .json regroupe les autres
        # fichiers du projet (plans, profil, style, sous-dossiers...).
        self.project_path = os.path.dirname(manifest_path)

        self._loaded_park = data.get("park_name", "")
        self._loaded_city = data.get("ref_city", "")
        self._loaded_year = data.get("ref_year", "")
        self._loaded_bureau = data.get("bureau_etude", "")

        self.lbl_project_status.setText(
            f"Projet chargé : {manifest_path}\n\n"
            f"Parc : {self._loaded_park or '-'}\n"
            f"Ville : {self._loaded_city or '-'}\n"
            f"Année : {self._loaded_year or '-'}\n"
            f"Bureau : {self._loaded_bureau or '-'}"
        )

    # ======================================================

    def try_accept(self):

        if self.mode == "open":

            if not self.project_path:
                QMessageBox.warning(
                    self, "Aucun projet choisi",
                    "Choisissez d'abord le fichier du projet à ouvrir."
                )
                return

            self.accept()
            return

        # Mode "create"
        if not self.chosen_location:
            QMessageBox.warning(
                self, "Emplacement manquant",
                "Choisissez l'emplacement où créer le projet."
            )
            return

        project_name = self.project_name_input.text().strip()

        if not project_name:
            project_name = "_".join(
                part for part in [
                    "Albat",
                    self.park_input.text().strip(),
                    self.year_input.text().strip()
                ] if part
            ) or "Albat_Projet"

        safe_name = "".join(
            c if c.isalnum() or c in " _-" else "_"
            for c in project_name
        ).strip() or "Albat_Projet"

        project_dir = os.path.join(self.chosen_location, safe_name)

        try:
            os.makedirs(project_dir, exist_ok=True)
            ensure_project_subfolders(project_dir)
        except Exception as e:
            QMessageBox.critical(
                self, "Erreur",
                f"Impossible de créer le dossier du projet :\n{e}"
            )
            return

        try:
            year_val = int(self.year_input.text().strip())
        except (ValueError, TypeError):
            year_val = _dt.date.today().year

        manifest = {
            "park_name": self.park_input.text().strip(),
            "ref_city": self.city_input.text().strip(),
            "ref_year": year_val,
            "bureau_etude": self.bureau_input.currentText().strip(),
        }

        try:
            with open(
                os.path.join(project_dir, f"{safe_name}.albatproj"),
                "w", encoding="utf-8"
            ) as f:
                json.dump(manifest, f, ensure_ascii=False, indent=2)
        except Exception as e:
            QMessageBox.critical(
                self, "Erreur",
                f"Impossible d'enregistrer le fichier de projet :\n{e}"
            )
            return

        self.project_path = project_dir

        self.accept()

    # ======================================================

    def get_values(self):

        if self.mode == "open" and self.project_path:
            try:
                year = int(self._loaded_year)
            except (ValueError, TypeError):
                year = _dt.date.today().year

            return (
                self._loaded_park,
                self._loaded_city,
                year,
                self._loaded_bureau,
                self.project_path,
                None
            )

        try:
            year = int(self.year_input.text().strip())
        except (ValueError, TypeError):
            year = _dt.date.today().year

        bureau_name = self.bureau_input.currentText().strip()
        bureau_profile_path = self.bureau_profiles.get(bureau_name)

        return (
            self.park_input.text().strip(),
            self.city_input.text().strip(),
            year,
            bureau_name,
            self.project_path,
            bureau_profile_path
        )


class Suite(QMainWindow):
    def __init__(self, splash=None):
        super().__init__()

        # Imports différés (lourds) : ils se font ici, une fois
        # l'écran de chargement affiché, plutôt qu'au démarrage
        # du script.
        from modules.Albat_Correlations_Fusion_Mac import BatCorrelations
        from modules.Albat_Graph_Fusion_Mac import GraphWindow
        from modules.Albat_Bridage_Fusion_Mac import BridageWindow
        from modules.Albat_Scenar_Mac import OptimisationWindow
        from modules.Albat_Rapport_Mac import RapportWindow

        self.setWindowTitle("Albat Fusion")

        # Taille adaptée à l'écran réel plutôt qu'une hauteur fixe
        # (1400px) pouvant dépasser la hauteur utile de l'écran —
        # la fenêtre se retrouvait alors partiellement hors champ,
        # sa partie basse inaccessible. On prend le plus petit
        # entre la taille "idéale" et l'espace réellement
        # disponible (barre des tâches déduite), avec une petite
        # marge, puis on centre la fenêtre sur l'écran.
        screen = QApplication.primaryScreen()
        available = (
            screen.availableGeometry() if screen else None
        )

        ideal_w, ideal_h = 700, 1400
        min_w, min_h = 620, 1040

        if available is not None:
            # Marge plus généreuse (80px, pas seulement 40) pour
            # laisser de la place à la barre de titre et aux bordures
            # de fenêtre, qui s'ajoutent APRÈS le resize() et ne sont
            # pas comptées dans available.height() : sans cette
            # marge, la fenêtre déborde légèrement en bas de l'écran
            # malgré un calcul de taille apparemment correct.
            margin = 80
            target_w = min(ideal_w, available.width() - margin)
            target_h = min(ideal_h, available.height() - margin)
            target_w = max(target_w, min(min_w, available.width()))
            target_h = max(target_h, min(min_h, available.height()))
        else:
            target_w, target_h = ideal_w, ideal_h

        self.resize(target_w, target_h)
        self.setMinimumSize(
            min(min_w, target_w), min(min_h, target_h)
        )

        if available is not None:
            # Centrage basé sur frameGeometry() (taille de la
            # fenêtre AVEC ses décorations : barre de titre,
            # bordures) plutôt qu'un calcul manuel sur la seule
            # taille du contenu — plus fiable, la barre de titre
            # étant uniquement en haut, un centrage qui l'ignore
            # pousse visuellement la fenêtre vers le bas.
            frame = self.frameGeometry()
            frame.moveCenter(available.center())
            self.move(frame.topLeft())

        from PySide6.QtGui import QIcon

        icon_path = resource_path("assets", "icone_albat.ico")

        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        self.setStyleSheet("""
        QMainWindow{
            background:#071107;
        }

        QToolTip{
            background-color:#1a1a1a;
            color:#eef7d0;
            border:1px solid rgba(150,184,99,180);
            padding:6px 10px;
            border-radius:6px;
            font-size:12px;
        }

        QTabWidget{
            background:transparent;
        }

        QTabWidget::pane{
            background:#071107;
            border:none;
        }

        QTabBar{
            background:transparent;
        }

        QTabWidget::tab-bar{
            alignment:center;
        }

        QTabBar::tab{
            background:rgba(90,127,71,140);
            color:#d8e6c0;
            min-width:100px;
            min-height:20px;
            margin:4px 4px 0px 4px;
            border-top-left-radius:12px;
            border-top-right-radius:12px;
            border-bottom-left-radius:0px;
            border-bottom-right-radius:0px;
            font-size:13px;
            font-weight:800;
            padding:8px 10px;
        }

        QTabBar::tab:selected{
            background:#6b9654;
            color:#ffffff;
            padding-top:10px;
            margin-top:2px;
        }

        QTabBar::tab:hover:!selected{
            background:#7da861;
        }
        """)

        # ==========================================
        # NOM DU PARC / VILLE DE RÉFÉRENCE (une seule fois)
        # ==========================================

        if splash is not None:
            splash.close()

        startup = StartupDialog()

        startup.exec()
        (
            self.park_name, self.ref_city, self.ref_year,
            self.bureau_etude, self.project_path,
            self.bureau_profile_path
        ) = startup.get_values()

        # S'assure que la structure de dossiers du projet existe
        # (utile aussi pour un ancien projet créé avant cette
        # fonctionnalité).
        dir_correlations = ""
        dir_graph = ""
        dir_bridage = ""
        dir_scenar = ""
        dir_rapport = ""
        dir_qgis = ""

        if self.project_path:
            ensure_project_subfolders(self.project_path)
            dir_correlations = os.path.join(self.project_path, "Correlations")
            dir_graph = os.path.join(self.project_path, "Graph")
            dir_bridage = os.path.join(self.project_path, "Bridage")
            dir_scenar = os.path.join(self.project_path, "Scenar")
            dir_rapport = os.path.join(self.project_path, "Rapport")
            dir_qgis = os.path.join(self.project_path, "Qgis")

        self.tabs = QTabWidget()
        # Le mode "document" (généralement pensé pour des onglets
        # façon navigateur) peut empêcher certains styles Qt de
        # respecter la propriété de centrage CSS
        # (QTabWidget::tab-bar{alignment:center}) : désactivé, sans
        # incidence visuelle puisque l'apparence des onglets est de
        # toute façon entièrement pilotée par la feuille de style
        # ci-dessous.
        self.tabs.setDocumentMode(False)
        self.tabs.setTabPosition(QTabWidget.North)
        # Qt dessine par défaut un fin trait ("base") sous la barre
        # d'onglets, au niveau du widget natif — indépendant de la
        # feuille de style CSS, donc "border:none" sur
        # QTabWidget::pane ne suffit pas toujours à le faire
        # disparaître selon le style actif (Fusion, ici). On le
        # désactive directement via l'API.
        self.tabs.tabBar().setDrawBase(False)
        self.tabs.tabBar().setExpanding(False)

        self.tab_correlations = BatCorrelations(
            park_name=self.park_name,
            ref_year=self.ref_year,
            bureau_etude=self.bureau_etude,
            default_save_dir=dir_correlations
        )
        self.tabs.addTab(self.tab_correlations, "CORRELATIONS")

        self.tab_graph = GraphWindow(
            park_name=self.park_name,
            ref_city=self.ref_city,
            ref_year=self.ref_year,
            bureau_etude=self.bureau_etude,
            default_save_dir=dir_graph,
            correlations_dir_hint=dir_correlations
        )
        self.tabs.addTab(self.tab_graph, "GRAPH")

        self.tab_bridage = BridageWindow(
            park_name=self.park_name,
            ref_city=self.ref_city,
            ref_year=self.ref_year,
            bureau_etude=self.bureau_etude,
            default_save_dir=dir_bridage,
            correlations_dir_hint=dir_correlations
        )
        self.tabs.addTab(self.tab_bridage, "BRIDAGE")

        self.tab_scenar = OptimisationWindow(
            park_name=self.park_name,
            ref_year=self.ref_year,
            ref_city=self.ref_city,
            bureau_etude=self.bureau_etude,
            default_save_dir=dir_scenar,
            correlations_dir_hint=dir_correlations
        )
        self.tabs.addTab(self.tab_scenar, "SCENAR")

        self.tab_rapport = RapportWindow(
            park_name=self.park_name,
            ref_city=self.ref_city,
            ref_year=self.ref_year,
            bureau_etude=self.bureau_etude,
            default_save_dir=dir_rapport,
            graph_dir_hint=dir_graph,
            bureau_profile_path=self.bureau_profile_path,
            scenar_dir_hint=dir_scenar,
            qgis_dir_hint=dir_qgis
        )
        self.tabs.addTab(self.tab_rapport, "RAPPORT")

        self.tabs.currentChanged.connect(self._on_main_tab_changed)


        # ==========================================
        # TOOLTIPS ONGLETS
        # ==========================================

        self.tabs.setTabToolTip(
            0,
            "Analyse les corrélations entre activité chiroptères et météo."
        )

        self.tabs.setTabToolTip(
            1,
            "Visualisation graphique et exploration des données."
        )

        self.tabs.setTabToolTip(
            2,
            "Génération des périodes de bridage éolien."
        )

        self.tabs.setTabToolTip(
            3,
            "Création de scénarios de bridage manuels ou automatiques."
        )

        self.tabs.setTabToolTip(
            4,
            "Génération du rapport type (étude d'impact)."
        )

        # ==========================================
        # CHARGEMENT D'UN PROJET (si sélectionné au popup)
        # ==========================================

        if self.project_path:
            self._load_project(self.project_path)

        # ==========================================
        # BOUTON "ENREGISTRER LE PROJET" (toujours visible)
        # ==========================================

        container = QWidget()
        container.setStyleSheet("background:#071107;")
        container_layout = QVBoxLayout(container)
        container_layout.setContentsMargins(8, 8, 8, 0)
        container_layout.setSpacing(6)

        save_project_btn = QPushButton("💾 Enregistrer le projet")
        save_project_btn.setStyleSheet("""
            QPushButton{
                background:rgba(59,41,25,220);
                color:white;
                border-radius:12px;
                padding:8px 16px;
                font-size:13px;
                font-weight:700;
            }
            QPushButton:hover{
                background:rgba(90,65,40,240);
            }
        """)
        save_project_btn.clicked.connect(self.save_project)

        self.autosave_status = QLabel("")
        self.autosave_status.setAlignment(Qt.AlignCenter)
        self.autosave_status.setStyleSheet("""
            background:transparent;
            color:rgba(238,247,208,140);
            font-size:10px;
            font-style:italic;
        """)

        container_layout.addWidget(save_project_btn)
        container_layout.addWidget(self.autosave_status)
        container_layout.addWidget(self.tabs)

        self.setCentralWidget(container)

        # Sauvegarde automatique discrète toutes les 2 minutes (voir
        # _auto_save_project) : ne fait rien tant qu'aucun projet
        # n'a été créé/ouvert au moins une fois manuellement.
        self.autosave_timer = QTimer(self)
        self.autosave_timer.setInterval(2 * 60 * 1000)
        self.autosave_timer.timeout.connect(self._auto_save_project)
        self.autosave_timer.start()

    # ======================================================

    def _on_main_tab_changed(self, index):
        """
        Rafraîchit automatiquement tous les dossiers/fichiers
        intégrés de l'onglet Rapport (matériel, QGIS, graphiques,
        Scenar, bibliographie) dès qu'on bascule dessus, pour ne
        pas avoir à cliquer sur chaque bouton "Rafraîchir"
        manuellement.
        """

        if self.tabs.widget(index) is self.tab_rapport:
            try:
                self.tab_rapport.refresh_all()
            except Exception:
                pass

    # ======================================================

    def _load_project(self, project_dir):

        errors = []

        bridage_path = os.path.join(project_dir, "bridage_plans.json")
        if os.path.exists(bridage_path):
            try:
                self.tab_bridage.load_plans_from_path(bridage_path)
            except Exception as e:
                errors.append(f"Plans Bridage : {e}")

        scenar_path = os.path.join(project_dir, "scenar_plan.json")
        if os.path.exists(scenar_path):
            try:
                self.tab_scenar.load_plan_from_path(scenar_path)
            except Exception as e:
                errors.append(f"Plan Scenar : {e}")

        rapport_profile_path = os.path.join(
            project_dir, "rapport_profile.json"
        )
        if os.path.exists(rapport_profile_path):
            try:
                self.tab_rapport.load_profile_from_path(
                    rapport_profile_path
                )
            except Exception as e:
                errors.append(f"Profil du bureau : {e}")

        rapport_style_path = os.path.join(
            project_dir, "rapport_style.json"
        )
        if os.path.exists(rapport_style_path):
            try:
                self.tab_rapport.load_style_from_path(rapport_style_path)
            except Exception as e:
                errors.append(f"Style du rapport : {e}")

        if errors:
            QMessageBox.warning(
                self, "Chargement partiel du projet",
                "Le projet a été chargé, mais certains éléments ont "
                "rencontré une erreur :\n\n" + "\n".join(errors)
            )

    # ======================================================

    def save_project(self):
        """
        Enregistre l'état de l'ensemble d'Albat Fusion (informations
        générales, plans Bridage, plan Scenar, profil et style du
        Rapport) dans un dossier de projet réutilisable.

        Si un projet est déjà chargé (ou déjà enregistré dans cette
        session), propose de l'écraser directement plutôt que de
        redemander systématiquement un dossier et un nom.
        """

        out_dir = None
        safe_name = None

        if self.project_path and os.path.isdir(self.project_path):

            msgbox = QMessageBox(self)
            msgbox.setWindowTitle("Enregistrer le projet")
            msgbox.setText(
                f"Un projet est déjà chargé :\n{self.project_path}\n\n"
                f"Voulez-vous écraser ce projet avec les "
                f"modifications actuelles, ou en créer un nouveau ?"
            )
            btn_overwrite = msgbox.addButton(
                "Écraser ce projet", QMessageBox.AcceptRole
            )
            btn_new = msgbox.addButton(
                "Nouveau projet", QMessageBox.DestructiveRole
            )
            msgbox.addButton("Annuler", QMessageBox.RejectRole)
            msgbox.exec()

            clicked = msgbox.clickedButton()

            if clicked is btn_overwrite:
                out_dir = self.project_path
                existing_manifests = [
                    f for f in os.listdir(out_dir)
                    if f.lower().endswith(".albatproj")
                ]
                safe_name = (
                    os.path.splitext(existing_manifests[0])[0]
                    if existing_manifests
                    else os.path.basename(out_dir.rstrip(os.sep))
                )
            elif clicked is not btn_new:
                # Bouton "Annuler" (ou fermeture de la fenêtre)
                return

        if out_dir is None:

            parent_dir = QFileDialog.getExistingDirectory(
                self, "Choisir le dossier où enregistrer le projet"
            )

            if not parent_dir:
                return

            default_name = "_".join(
                part for part in [
                    "Albat",
                    self.park_name or "",
                    str(self.ref_year) if self.ref_year else ""
                ] if part
            )

            project_name, ok = QInputDialog.getText(
                self, "Nom du projet",
                "Nom du projet :",
                text=default_name
            )

            if not ok or not project_name.strip():
                return

            safe_name = "".join(
                c if c.isalnum() or c in " _-" else "_"
                for c in project_name.strip()
            ).strip() or default_name

            out_dir = os.path.join(parent_dir, safe_name)

        # Le dossier choisi/tapé dans la fenêtre peut ne pas encore
        # exister réellement sur le disque (ex : "Nouveau dossier"
        # créé dans la barre d'adresse sans validation) — on le crée
        # au besoin avant d'y écrire quoi que ce soit.
        try:
            os.makedirs(out_dir, exist_ok=True)
            ensure_project_subfolders(out_dir)
        except Exception as e:
            QMessageBox.critical(
                self, "Erreur",
                f"Impossible de créer le dossier :\n{out_dir}\n\n{e}"
            )
            return

        self._write_project_files(out_dir, safe_name, silent=False)

    # ======================================================

    def _write_project_files(self, out_dir, safe_name, silent=False):
        """
        Écrit effectivement tous les fichiers du projet dans
        out_dir (manifeste, plans Bridage/Scenar, profil et style
        du Rapport) et relie chaque onglet à son sous-dossier.
        Factorisé hors de save_project() pour être réutilisé tel
        quel par la sauvegarde automatique discrète (silent=True :
        aucune boîte de dialogue, juste un petit indicateur textuel
        mis à jour — voir _auto_save_project).
        """

        saved = []
        errors = []

        manifest_filename = f"{safe_name}.albatproj"

        try:
            manifest = {
                "park_name": self.park_name,
                "ref_city": self.ref_city,
                "ref_year": self.ref_year,
                "bureau_etude": self.bureau_etude,
            }
            with open(
                os.path.join(out_dir, manifest_filename),
                "w", encoding="utf-8"
            ) as f:
                json.dump(manifest, f, ensure_ascii=False, indent=2)
            saved.append("Informations générales")
        except Exception as e:
            errors.append(f"Informations générales : {e}")

        try:
            self.tab_bridage.save_plans_to_path(
                os.path.join(out_dir, "bridage_plans.json")
            )
            saved.append("Plans Bridage")
        except Exception as e:
            errors.append(f"Plans Bridage : {e}")

        try:
            self.tab_scenar.save_plan_to_path(
                os.path.join(out_dir, "scenar_plan.json")
            )
            saved.append("Plan Scenar")
        except Exception as e:
            errors.append(f"Plan Scenar : {e}")

        try:
            self.tab_rapport.save_profile_to_path(
                os.path.join(out_dir, "rapport_profile.json")
            )
            saved.append("Profil du bureau (Rapport)")
        except Exception as e:
            errors.append(f"Profil du bureau : {e}")

        try:
            self.tab_rapport.save_style_to_path(
                os.path.join(out_dir, "rapport_style.json")
            )
            saved.append("Style du rapport")
        except Exception as e:
            errors.append(f"Style du rapport : {e}")

        self.project_path = out_dir

        # Relie chaque onglet à son sous-dossier, utile si le projet
        # vient d'être créé en cours de session (pas via le popup
        # de démarrage) : les prochains enregistrements de fichiers
        # depuis les onglets cibleront directement le bon dossier.
        self.tab_correlations.default_save_dir = os.path.join(
            out_dir, "Correlations"
        )
        self.tab_graph.default_save_dir = os.path.join(out_dir, "Graph")
        self.tab_bridage.default_save_dir = os.path.join(out_dir, "Bridage")
        self.tab_scenar.default_save_dir = os.path.join(out_dir, "Scenar")
        self.tab_rapport.default_save_dir = os.path.join(out_dir, "Rapport")

        if not self.tab_rapport.graphs_dir:
            self.tab_rapport.graphs_dir = os.path.join(out_dir, "Graph")

        if not self.tab_rapport.qgis_dir:
            self.tab_rapport.qgis_dir = os.path.join(out_dir, "Qgis")

        if silent:
            # Sauvegarde automatique : pas de fenêtre, juste un
            # petit indicateur textuel discret sous le bouton
            # "Enregistrer le projet", qui s'efface tout seul après
            # quelques secondes.
            import datetime as _dtmod
            heure = _dtmod.datetime.now().strftime("%H:%M")
            if errors and not saved:
                return
            self.autosave_status.setText(
                f"Sauvegarde automatique effectuée à {heure}"
            )
            QTimer.singleShot(
                8000, lambda: self.autosave_status.setText("")
            )
            return

        msg = (
            f"Projet enregistré dans :\n{out_dir}\n\n"
            "Éléments sauvegardés :\n"
            + "\n".join(f"✓ {s}" for s in saved)
        )

        if errors:
            msg += (
                "\n\nNon sauvegardés (facultatifs, souvent car pas "
                "encore renseignés) :\n"
                + "\n".join(f"✗ {e}" for e in errors)
            )

        QMessageBox.information(self, "Enregistrement du projet", msg)

    # ======================================================

    def _auto_save_project(self):
        """
        Sauvegarde automatique discrète, appelée toutes les 2
        minutes par un QTimer (voir __init__). Ne fait rien tant
        qu'aucun projet n'a été créé/ouvert (pas de dossier connu
        où écrire) — dans ce cas, l'utilisateur doit toujours faire
        un premier "Enregistrer le projet" manuel pour choisir
        l'emplacement. Erreurs volontairement avalées : une
        sauvegarde automatique ratée ne doit jamais interrompre le
        travail en cours avec une fenêtre d'erreur.
        """

        if not self.project_path or not os.path.isdir(self.project_path):
            return

        try:
            existing_manifests = [
                f for f in os.listdir(self.project_path)
                if f.lower().endswith(".albatproj")
            ]
            safe_name = (
                os.path.splitext(existing_manifests[0])[0]
                if existing_manifests
                else os.path.basename(self.project_path.rstrip(os.sep))
            )
            self._write_project_files(
                self.project_path, safe_name, silent=True
            )
        except Exception:
            pass


if __name__ == "__main__":

    # Masque la fenêtre console (invite de commandes) sur Windows,
    # si le script a été lancé avec python.exe plutôt que pythonw.exe.
    if sys.platform.startswith("win"):
        try:
            import ctypes
            hwnd = ctypes.windll.kernel32.GetConsoleWindow()
            if hwnd:
                ctypes.windll.user32.ShowWindow(hwnd, 0)
        except Exception:
            pass

        # Déclare une identité propre à l'application auprès de
        # Windows (Application User Model ID), impérativement AVANT
        # toute création de fenêtre. Sans ça, Windows associe le
        # processus à l'icône générique de Python dans la barre des
        # tâches, même si setWindowIcon() est bien appelé sur
        # chaque fenêtre — l'icône ne s'affiche alors correctement
        # que dans la barre de titre, pas dans la barre des tâches
        # ni dans Alt+Tab.
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                "SilvaEnvironnement.AlbatFusion.Suite.1"
            )
        except Exception:
            pass

    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    # Icône au niveau de l'application (en plus de chaque fenêtre
    # individuelle) : couvre aussi les boîtes de dialogue et
    # renforce la cohérence de l'icône affichée par Windows/macOS.
    from PySide6.QtGui import QIcon
    icon_path = resource_path("assets", "icone_albat.ico")
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))

    # Écran de chargement avec le logo, pendant l'initialisation
    # (imports des modules + construction de la fenêtre principale).
    splash = None
    splash_path = resource_path("assets", "chauve_souris_transparent.png")

    if os.path.exists(splash_path):
        splash_pix = QPixmap(splash_path)
        if not splash_pix.isNull():
            splash_pix = splash_pix.scaled(
                340, 340,
                Qt.KeepAspectRatio,
                Qt.SmoothTransformation
            )
            splash = QSplashScreen(splash_pix)
            splash.setAttribute(Qt.WA_TranslucentBackground)
            splash.setWindowFlag(Qt.FramelessWindowHint)
            splash.showMessage(
                "Chargement d'Albat Fusion...",
                Qt.AlignBottom | Qt.AlignHCenter,
                Qt.white
            )
            splash.show()
            app.processEvents()

    win = Suite(splash=splash)

    # Centrer la fenêtre au milieu de l'écran
    screen = app.primaryScreen().availableGeometry()
    win_geo = win.frameGeometry()
    win_geo.moveCenter(screen.center())
    win.move(win_geo.topLeft())

    win.show()

    sys.exit(app.exec())