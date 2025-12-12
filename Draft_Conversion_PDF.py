import tkinter as tk 
from tkinter import filedialog, messagebox, ttk
import os
import threading
from pathlib import Path

# ============================================================================
# IMPORTS DES MODULES DE CONVERSION
# ============================================================================

try:
    import win32com.client
    WINDOWS_COM_AVAILABLE = True
except ImportError:
    WINDOWS_COM_AVAILABLE = False

try:
    from docx2pdf import convert as docx_convert
    DOCX2PDF_AVAILABLE = True
except ImportError:
    DOCX2PDF_AVAILABLE = False


# ============================================================================
# CLASSE PRINCIPALE
# ============================================================================

class PDFConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Convertisseur Word/Excel vers PDF")
        self.root.geometry("800x600")

        self.selected_folder = ""
        self.files_list = []
        self.conversion_in_progress = False
        self.include_subfolders = tk.BooleanVar(value=True)  # Option pour les sous-dossiers

        self.setup_ui()
        self.check_dependencies()

    # ========================================================================
    # 1. CONFIGURATION DE L'INTERFACE GRAPHIQUE
    # ========================================================================

    def setup_ui(self):
        """Configure tous les éléments de l'interface utilisateur"""
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configuration du grid
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)

        # --- Section sélection du dossier ---
        ttk.Label(main_frame, text="Dossier sélectionné :").grid(
            row=0, column=0, sticky=tk.W, pady=(0, 5))

        folder_frame = ttk.Frame(main_frame)
        folder_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        folder_frame.columnconfigure(0, weight=1)

        self.folder_label = ttk.Label(
            folder_frame, text="Aucun dossier sélectionné",
            background="white", relief="sunken")
        self.folder_label.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))

        self.browse_button = ttk.Button(
            folder_frame, text="Parcourir", command=self.browse_folder)
        self.browse_button.grid(row=0, column=1)

        # --- Option inclure les sous-dossiers ---
        subfolder_frame = ttk.Frame(main_frame)
        subfolder_frame.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=(0, 5))
        
        self.subfolder_checkbox = ttk.Checkbutton(
            subfolder_frame, 
            text="Inclure les fichiers des sous-dossiers",
            variable=self.include_subfolders,
            command=self.list_files
        )
        self.subfolder_checkbox.grid(row=0, column=0, sticky=tk.W)

        # --- Section liste des fichiers ---
        files_frame = ttk.LabelFrame(main_frame, text="Fichiers trouvés", padding="5")
        files_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        files_frame.columnconfigure(0, weight=1)
        files_frame.rowconfigure(0, weight=1)

        columns = ("Type", "Nom", "Taille", "Statut")
        self.files_tree = ttk.Treeview(files_frame, columns=columns, show="tree headings")

        self.files_tree.heading("#0", text="Sélection")
        self.files_tree.heading("Type", text="Type")
        self.files_tree.heading("Nom", text="Nom du fichier")
        self.files_tree.heading("Taille", text="Taille")
        self.files_tree.heading("Statut", text="Statut")

        self.files_tree.column("#0", width=80, minwidth=80)
        self.files_tree.column("Type", width=80, minwidth=60)
        self.files_tree.column("Nom", width=450, minwidth=300)
        self.files_tree.column("Taille", width=100, minwidth=80)
        self.files_tree.column("Statut", width=130, minwidth=100)

        scrollbar = ttk.Scrollbar(files_frame, orient="vertical", command=self.files_tree.yview)
        self.files_tree.configure(yscrollcommand=scrollbar.set)

        self.files_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

        # --- Boutons de sélection ---
        selection_frame = ttk.Frame(main_frame)
        selection_frame.grid(row=4, column=0, columnspan=2, pady=(0, 10))

        ttk.Button(selection_frame, text="Tout sélectionner",
                   command=self.select_all).grid(row=0, column=0, padx=(0, 5))
        ttk.Button(selection_frame, text="Tout désélectionner",
                   command=self.deselect_all).grid(row=0, column=1, padx=(0, 5))
        ttk.Button(selection_frame, text="Inverser la sélection",
                   command=self.invert_selection).grid(row=0, column=2)

        # --- Bouton de conversion ---
        self.convert_button = ttk.Button(
            main_frame, text="Convertir en PDF",
            command=self.start_conversion, state="disabled")
        self.convert_button.grid(row=5, column=0, columnspan=2, pady=(0, 10))

        # --- Barre de progression ---
        progress_frame = ttk.LabelFrame(main_frame, text="Progression", padding="5")
        progress_frame.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        progress_frame.columnconfigure(0, weight=1)

        self.progress = ttk.Progressbar(progress_frame, mode="determinate")
        self.progress.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))

        self.status_label = ttk.Label(progress_frame, text="Prêt")
        self.status_label.grid(row=1, column=0, sticky=tk.W)

    # ========================================================================
    # 2. VÉRIFICATION DES DÉPENDANCES
    # ========================================================================

    def check_dependencies(self):
        """Vérifie que les modules nécessaires sont installés"""
        missing_deps = []

        if not WINDOWS_COM_AVAILABLE:
            missing_deps.append("pywin32 (pour la conversion Office)")

        if not DOCX2PDF_AVAILABLE:
            missing_deps.append("docx2pdf (alternative pour Word)")

        if missing_deps:
            msg = "Dépendances manquantes :\n" + "\n".join(f"- {dep}" for dep in missing_deps)
            msg += "\n\nInstallez-les avec :\npip install pywin32 docx2pdf"
            messagebox.showwarning("Dépendances manquantes", msg)

    # ========================================================================
    # 3. GESTION DES FICHIERS
    # ========================================================================

    def browse_folder(self):
        """Ouvre un dialogue pour sélectionner un dossier avec aperçu des fichiers"""
        # Utilise le dialogue natif qui montre tous les types de fichiers
        folder = filedialog.askdirectory(
            title="Sélectionnez un dossier (tous les fichiers sont visibles dans l'explorateur)",
            mustexist=True,
            initialdir=self.selected_folder if self.selected_folder else None
        )
        if folder:
            self.selected_folder = folder
            self.folder_label.config(text=folder)
            self.list_files()

    def list_files(self):
        """Liste tous les fichiers Word et Excel dans le dossier sélectionné"""
        if not self.selected_folder:
            return

        # Extensions supportées avec leur type
        extensions = {
            '.docx': 'Word',
            '.doc': 'Word',
            '.xlsx': 'Excel',
            '.xls': 'Excel',
            '.xlsm': 'Excel'
        }

        # Vide la liste actuelle
        for item in self.files_tree.get_children():
            self.files_tree.delete(item)

        self.files_list = []

        try:
            # Choisit entre recherche récursive ou seulement le dossier principal
            if self.include_subfolders.get():
                file_iterator = Path(self.selected_folder).rglob("*")
            else:
                file_iterator = Path(self.selected_folder).glob("*")
            
            for file_path in file_iterator:
                if file_path.is_file() and file_path.suffix.lower() in extensions:
                    file_type = extensions[file_path.suffix.lower()]
                    
                    file_info = {
                        'path': str(file_path),
                        'name': file_path.name,
                        'type': file_type,
                        'size': self.format_size(file_path.stat().st_size),
                        'selected': False,
                        'status': 'En attente'
                    }
                    self.files_list.append(file_info)

                    # Ajoute à la treeview
                    self.files_tree.insert(
                        "", "end", text="☐",
                        values=(file_info['type'], file_info['name'], 
                                file_info['size'], file_info['status']))

            # Active le bouton si des fichiers sont trouvés
            if self.files_list:
                self.convert_button.config(state="normal")
                word_count = sum(1 for f in self.files_list if f['type'] == 'Word')
                excel_count = sum(1 for f in self.files_list if f['type'] == 'Excel')
                status_text = f"{len(self.files_list)} fichier(s) trouvé(s) - Word: {word_count}, Excel: {excel_count}"
                self.status_label.config(text=status_text)
            else:
                self.convert_button.config(state="disabled")
                self.status_label.config(text="Aucun fichier Word/Excel trouvé")

            # Lie les événements de clic
            self.files_tree.bind("<Button-1>", self.on_tree_click)

        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de la lecture du dossier :\n{str(e)}")

    def format_size(self, size_bytes):
        """Formate la taille d'un fichier de manière lisible"""
        if size_bytes < 1024:
            return f"{size_bytes} B"
        elif size_bytes < 1024 ** 2:
            return f"{size_bytes / 1024:.1f} KB"
        else:
            return f"{size_bytes / (1024 ** 2):.1f} MB"

    # ========================================================================
    # 4. GESTION DE LA SÉLECTION
    # ========================================================================

    def on_tree_click(self, event):
        """Gère les clics sur la liste pour sélectionner/désélectionner"""
        item = self.files_tree.identify('item', event.x, event.y)
        if item:
            children = self.files_tree.get_children()
            try:
                index = list(children).index(item)
                if 0 <= index < len(self.files_list):
                    self.files_list[index]['selected'] = not self.files_list[index]['selected']
                    new_text = "☑" if self.files_list[index]['selected'] else "☐"
                    self.files_tree.item(item, text=new_text)
            except (ValueError, IndexError):
                pass

    def select_all(self):
        """Sélectionne tous les fichiers"""
        for i, file_info in enumerate(self.files_list):
            file_info['selected'] = True
            item = self.files_tree.get_children()[i]
            self.files_tree.item(item, text="☑")

    def deselect_all(self):
        """Désélectionne tous les fichiers"""
        for i, file_info in enumerate(self.files_list):
            file_info['selected'] = False
            item = self.files_tree.get_children()[i]
            self.files_tree.item(item, text="☐")

    def invert_selection(self):
        """Inverse la sélection actuelle"""
        for i, file_info in enumerate(self.files_list):
            file_info['selected'] = not file_info['selected']
            new_text = "☑" if file_info['selected'] else "☐"
            item = self.files_tree.get_children()[i]
            self.files_tree.item(item, text=new_text)

    # ========================================================================
    # 5. CONVERSION DES FICHIERS
    # ========================================================================

    def start_conversion(self):
        """Démarre le processus de conversion dans un thread séparé"""
        if self.conversion_in_progress:
            return

        selected_files = [f for f in self.files_list if f['selected']]
        if not selected_files:
            messagebox.showwarning("Aucune sélection",
                                   "Veuillez sélectionner au moins un fichier.")
            return

        self.conversion_in_progress = True
        self.convert_button.config(state="disabled")

        thread = threading.Thread(target=self.convert_files, args=(selected_files,))
        thread.daemon = True
        thread.start()

    def convert_files(self, selected_files):
        """Convertit les fichiers sélectionnés en PDF"""
        total_files = len(selected_files)
        success_count = 0
        error_count = 0
        old_pdfs = []  # Liste des PDFs renommés en "_ancien"

        for i, file_info in enumerate(selected_files):
            # Met à jour le statut : En cours
            self.update_file_status(file_info, "En cours...")
            
            try:
                # Met à jour la progression
                progress = (i / total_files) * 100
                self.root.after(0, lambda p=progress: self.progress.config(value=p))
                self.root.after(0, lambda name=file_info['name'], ftype=file_info['type']:
                self.status_label.config(text=f"Conversion {ftype}: {name}"))

                # Détermine le fichier de sortie
                input_path = Path(file_info['path'])
                output_path = input_path.with_suffix('.pdf')

                # Renomme le PDF existant en ajoutant "_ancien" avant de créer le nouveau
                if output_path.exists():
                    try:
                        # Crée le nouveau nom avec "_ancien"
                        old_pdf_name = output_path.stem + "_ancien.pdf"
                        old_pdf_path = output_path.parent / old_pdf_name
                        
                        # Si un fichier "_ancien.pdf" existe déjà, le supprime
                        if old_pdf_path.exists():
                            old_pdf_path.unlink()
                        
                        # Renomme le PDF actuel en "_ancien"
                        output_path.rename(old_pdf_path)
                        old_pdfs.append(old_pdf_path)  # Ajoute à la liste pour archivage
                    except Exception as e:
                        print(f"Impossible de renommer {output_path}: {str(e)}")

                # Convertit le fichier selon son type
                if file_info['type'] == 'Word':
                    self.convert_word_to_pdf(str(input_path), str(output_path))
                elif file_info['type'] == 'Excel':
                    self.convert_excel_to_pdf(str(input_path), str(output_path))
                
                # Met à jour le statut : Terminé
                self.update_file_status(file_info, "✓ Terminé")
                success_count += 1

            except Exception as e:
                # Met à jour le statut : Erreur
                self.update_file_status(file_info, "✗ Erreur")
                error_count += 1
                print(f"Erreur lors de la conversion de {file_info['name']}: {str(e)}")

        # Finalise la conversion
        self.root.after(0, lambda: self.progress.config(value=100))
        self.root.after(0, lambda: self.status_label.config(
            text=f"Terminé : {success_count} réussi(s), {error_count} erreur(s)"))
        self.root.after(0, lambda: self.convert_button.config(state="normal"))

        # Déplace tous les fichiers "_ancien.pdf" dans le dossier pdf_archive
        if old_pdfs:
            self.archive_old_pdfs(old_pdfs)

        self.conversion_in_progress = False

        # Affiche un message de fin
        if error_count == 0:
            archive_msg = f"\n{len(old_pdfs)} ancien(s) PDF archivé(s)" if old_pdfs else ""
            self.root.after(0, lambda: messagebox.showinfo(
                "Conversion terminée",
                f"Tous les fichiers ont été convertis avec succès !{archive_msg}"))
        else:
            self.root.after(0, lambda: messagebox.showwarning(
                "Conversion terminée",
                f"Conversion terminée avec {error_count} erreur(s)."))

    def archive_old_pdfs(self, old_pdfs):
        """Déplace tous les fichiers PDF anciens dans un dossier pdf_archive"""
        if not old_pdfs:
            return
        
        try:
            # Crée le dossier pdf_archive dans le dossier sélectionné
            archive_folder = Path(self.selected_folder) / "pdf_archive"
            archive_folder.mkdir(exist_ok=True)
            
            # Déplace chaque fichier ancien dans le dossier d'archive
            for old_pdf_path in old_pdfs:
                if old_pdf_path.exists():
                    destination = archive_folder / old_pdf_path.name
                    
                    # Si le fichier existe déjà dans l'archive, le supprime
                    if destination.exists():
                        destination.unlink()
                    
                    # Déplace le fichier
                    old_pdf_path.rename(destination)
            
            print(f"{len(old_pdfs)} fichier(s) archivé(s) dans {archive_folder}")
        
        except Exception as e:
            print(f"Erreur lors de l'archivage des anciens PDFs : {str(e)}")

    def update_file_status(self, file_info, status):
        """Met à jour le statut d'un fichier dans la treeview"""
        file_info['status'] = status
        
        def update():
            # Trouve l'item correspondant dans la treeview
            for i, f in enumerate(self.files_list):
                if f['path'] == file_info['path']:
                    item = self.files_tree.get_children()[i]
                    # Met à jour seulement la colonne statut
                    current_values = list(self.files_tree.item(item, 'values'))
                    current_values[-1] = status  # Dernière colonne = Statut
                    self.files_tree.item(item, values=current_values)
                    break
        
        self.root.after(0, update)

    def convert_word_to_pdf(self, input_path, output_path):
        """Convertit un fichier Word en PDF"""
        if WINDOWS_COM_AVAILABLE:
            # Méthode 1 : COM (recommandé sur Windows)
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            try:
                doc = word.Documents.Open(input_path)
                doc.SaveAs(output_path, FileFormat=17)  # 17 = PDF
                doc.Close()
            finally:
                word.Quit()
        elif DOCX2PDF_AVAILABLE:
            # Méthode 2 : docx2pdf (alternative)
            docx_convert(input_path, output_path)
        else:
            raise Exception("Aucun module de conversion Word disponible")

    def convert_excel_to_pdf(self, input_path, output_path):
        """Convertit un fichier Excel en PDF (TOUTES LES FEUILLES) - EN ARRIÈRE-PLAN"""
        if not WINDOWS_COM_AVAILABLE:
            raise Exception("pywin32 est requis pour convertir Excel en PDF")
        
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False  # Mode arrière-plan
        excel.DisplayAlerts = False  # Pas de pop-ups
        excel.ScreenUpdating = False  # Pas de rafraîchissement d'écran
        
        try:
            # Ouvre le classeur en arrière-plan
            workbook = excel.Workbooks.Open(input_path, ReadOnly=True, UpdateLinks=False)
            
            # Exporte tout le classeur en PDF
            workbook.ExportAsFixedFormat(
                Type=0,  # 0 = PDF
                Filename=output_path,
                Quality=0,  # 0 = Standard quality
                IncludeDocProperties=True,
                IgnorePrintAreas=False,  # Respecte les zones d'impression définies
                OpenAfterPublish=False
            )
            
            workbook.Close(SaveChanges=False)
            
        finally:
            excel.ScreenUpdating = True
            excel.Quit()


# ============================================================================
# POINT D'ENTRÉE
# ============================================================================

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFConverter(root)
    root.mainloop() 
