from openpyxl import load_workbook
from datetime import datetime

def write_to_excel(path, ZLME_data, EURX_data, log_function):
    try:
        # Charger le fichier Excel existant
        book = load_workbook(path)
    except Exception as e:
        log_function(f'Erreur lors du chargement du fichier Excel : {e}')
        return

    # Assurez-vous que les données sont triées par date
    ZLME_data_sorted = sorted(ZLME_data, key=lambda x: datetime.strptime(x[0], '%d.%m.%Y'))
    EURX_data_sorted = sorted(EURX_data, key=lambda x: datetime.strptime(x[0], '%d.%m.%Y'))

    # Écrire les données ZLME
    try:
        sheet = book['ZLME']
        for data in ZLME_data_sorted:
            # Vérifiez que chaque donnée est un tuple avec au moins 2 éléments (date et valeur)
            if isinstance(data, tuple) and len(data) >= 2:
                # Convertir la date en objet datetime et la formater comme vous le souhaitez
                date = datetime.strptime(data[0], '%d.%m.%Y').strftime('%d.%m.%Y')
                value = data[1]
                new_row = ['ZLME', date, value, 'USD', 'EUR']
                sheet.append(new_row)
                log_function(new_row)
    except Exception as e:
        log_function(f'Erreur lors de l\'écriture dans la feuille ZLME : {e}')
        return

    # Écrire les données EURX
    try:
        sheet = book['EURX']
        for data in EURX_data_sorted:
            if isinstance(data, tuple) and len(data) >= 2:
                date = datetime.strptime(data[0], '%d.%m.%Y').strftime('%d.%m.%Y')
                value = data[1]
                new_row = ['EURX', date, value, 'USD', 'EUR']
                sheet.append(new_row)
                log_function(new_row)
    except Exception as e:
        log_function(f'Erreur lors de l\'écriture dans la feuille EURX : {e}')
        return

    # Sauvegarder le fichier Excel
    try:
        book.save(path)
        log_function("Les données ont été ajoutées avec succès dans le fichier Excel !")
    except Exception as e:
        log_function(f"Erreur lors de la sauvegarde du fichier Excel : {e}")
