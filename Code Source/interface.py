from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QFileDialog, QVBoxLayout,
    QLabel, QProgressBar, QHBoxLayout, QRadioButton, QFrame
)
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5.QtGui import QPixmap, QIcon, QFont, QColor, QPalette
import sys, os

from comparaison import comparaisonAzimut, comparaisonCoordonee  # Fonctions de traitement à importer


# Thread pour exécuter la comparaison sans bloquer l'interface
class ComparaisonThread(QThread):
    progress = pyqtSignal(int)      # Signal de progression
    finished = pyqtSignal(str)      # Signal envoyé avec le chemin du fichier généré

    def __init__(self, file_path, output_path, mode):
        """
        Initialise le thread de comparaison.
        """
        super().__init__()
        self.file_path = file_path
        self.output_path = output_path
        self.mode = mode

    def run(self):
        """ Exécute la comparaison dans un thread séparé."""
        # Lancer la bonne fonction selon le mode sélectionné
        if self.mode == "azimut":
            comparaisonAzimut(self.file_path, self.output_path, self.progress.emit)
        else:
            comparaisonCoordonee(self.file_path, self.output_path, self.progress.emit)
        self.finished.emit(self.output_path)



# Interface principale
class Interface(QWidget):

    """
    Interface graphique PyQt5 pour la comparaison d'azimuts ou de coordonnées dans un fichier Excel.
    L'utilisateur peut sélectionner un fichier Excel et choisir entre deux types de comparaison :
    - Comparaison des azimuts de rayonnement (2G, 3G, 4G)
    - Comparaison des coordonnées géographiques (X, Y)

    Les résultats sont exportés dans un fichier Excel de sortie avec une colonne supplémentaire.
    """

    def __init__(self):
        """ Initialisation de l'interface graphique. """
        super().__init__()
        self.setWindowTitle("ComparateurRadioTT")  # Titre de la fenêtre
        self.setWindowIcon(QIcon("icon.ico"))  # Icône personnalisée
        self.resize(800, 500)
        self.setup_ui()       # Création des composants
        self.apply_styles()   # Application des styles

    def setup_ui(self):
        """ Configuration de l'interface utilisateur. """
        # Layout principal avec marges
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(40, 30, 40, 30)
        main_layout.setSpacing(20)

        # Logo (en haut)
        header = QHBoxLayout()
        self.logo_label = QLabel()
        self.logo_label.setPixmap(QPixmap("logo.png").scaledToHeight(120, Qt.SmoothTransformation))
        header.addWidget(self.logo_label, alignment=Qt.AlignCenter)
        main_layout.addLayout(header)

        # Carte contenant tous les éléments
        card = QFrame()
        card.setObjectName("mainCard")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(30, 30, 30, 30)
        card_layout.setSpacing(25)

        # Titre principal
        self.title_label = QLabel("Comparateur Radio TT")
        self.title_label.setObjectName("titleLabel")
        card_layout.addWidget(self.title_label, alignment=Qt.AlignCenter)

        # Description
        self.desc_label = QLabel("Comparer les paramètres radio ou les coordonnées dans un fichier Excel")
        self.desc_label.setObjectName("descLabel")
        card_layout.addWidget(self.desc_label, alignment=Qt.AlignCenter)

        # Options de comparaison (azimut ou coordonnées)
        options_frame = QFrame()
        options_frame.setObjectName("optionsFrame")
        options_layout = QVBoxLayout(options_frame)
        options_layout.setContentsMargins(20, 20, 20, 20)
        options_layout.setSpacing(15)

        self.radio_azimut = QRadioButton("Comparaison des azimuts")
        self.radio_coord = QRadioButton("Comparaison des coordonnées")
        self.radio_azimut.setChecked(True)

        self.radio_azimut.setObjectName("azimutRadio")
        self.radio_coord.setObjectName("coordRadio")

        options_layout.addWidget(self.radio_azimut)
        options_layout.addWidget(self.radio_coord)
        card_layout.addWidget(options_frame)

        # Bouton de sélection de fichier
        self.action_btn = QPushButton("SÉLECTIONNER UN FICHIER EXCEL")
        self.action_btn.setObjectName("actionButton")
        self.action_btn.setCursor(Qt.PointingHandCursor)
        self.action_btn.setMinimumHeight(50)
        self.action_btn.clicked.connect(self.choisir_fichier)
        card_layout.addWidget(self.action_btn)

        # Barre de progression
        self.progress_bar = QProgressBar()
        self.progress_bar.setObjectName("progressBar")
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setMinimumHeight(8)
        card_layout.addWidget(self.progress_bar)

        # Étiquette de statut
        self.status_label = QLabel("Prêt à analyser")
        self.status_label.setObjectName("statusLabel")
        card_layout.addWidget(self.status_label, alignment=Qt.AlignCenter)

        main_layout.addWidget(card)
        self.setLayout(main_layout)

    def load_stylesheet(self, path):
        """ Charge une feuille de style CSS depuis un fichier.
        Args:
            path (str): Chemin vers le fichier CSS.
        Returns:
            str: Contenu de la feuille de style CSS.
        """
        with open(path, "r") as f:
            return f.read()

    def apply_styles(self):
        """ Applique les styles à l'interface graphique. """
        # Thème clair (fond blanc et texte sombre)
        palette = self.palette()
        palette.setColor(QPalette.Window, QColor(248, 250, 252))
        palette.setColor(QPalette.WindowText, QColor(30, 41, 59))
        self.setPalette(palette)

        self.setStyleSheet(self.load_stylesheet("style.css"))# Feuille de style CSS
        

    def choisir_fichier(self):
        """ Ouvre un dialogue pour sélectionner un fichier Excel et lance le traitement. """
        # Ouvrir un explorateur de fichiers
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Sélectionner un fichier Excel",
            "",
            "Fichiers Excel (*.xlsx);;Tous les fichiers (*)"
        )

        if file_path:
            output_path = file_path.replace(".xlsx", "_result.xlsx")
            mode = "azimut" if self.radio_azimut.isChecked() else "coord"

            # Préparer l'affichage
            self.status_label.setText("Traitement en cours...")
            self.progress_bar.setValue(0)
            self.action_btn.setEnabled(False)

            # Lancer le traitement dans un thread séparé
            self.thread = ComparaisonThread(file_path, output_path, mode)
            self.thread.progress.connect(self.progress_bar.setValue)
            self.thread.finished.connect(self.processing_complete)
            self.thread.start()

    def processing_complete(self, output_path):
        """ Mise à jour de l'interface après la fin du traitement. """
        self.status_label.setText(f"Analyse terminée : {os.path.basename(output_path)}")
        self.progress_bar.setValue(100)
        self.action_btn.setEnabled(True)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setFont(QFont("Segoe UI", 10))
    window = Interface()
    window.show()
    sys.exit(app.exec_())
