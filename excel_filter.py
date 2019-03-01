import pandas as pd


'''
Funkcja filtruje podaną kolumnę pliku Excel na podstawie niechcianych danych z drugiej podanej kolumny 
w innym pliku Excel'a. Po filtrowaniu dane są zapisywane do nowego pliku. Do poprawnego działania pliki
do edycji oraz skrypt powinny znajdować się w ten samej lokalizacji. W razie potrzeby można dodać ścieżkę bezwględną. 
'''

source_file_path = 'source.xlsx' # ściezka do pliku bazowego
source_file_column_to_be_filtered_name = 'kolumna' # nazwa kolumny z arkusza źródłowego, która zostanie poddana filtrowaniu

unwanted_phrases_file_path = 'unwanted_data.xlsx' # ścieżka do pliku zawierającego niechciane dane, które chcemy odfiltrować
unwanted_phrases_column_name = 'sheet_name' # nazwa kolumny zawierającej nieporządane dane

filtered_file_path = 'resoult.xlsx' # ścieżka do pliku wynikowego (jeżeli nie istnieje, skrypt utworzy go sam [o ile posiada uprawnienia do tworzenia plików])

def excel_editor(source_file_path, unwanted_phrases_file_path, filtered_file_path, source_file_column_to_be_filtered_name, unwanted_phrases_column_name ):
    source_file_data_frame = pd.read_excel(source_file_path)
    unwanted_phrases_data_frame = pd.read_excel(unwanted_phrases_file_path)
    filtered_column = (row[source_file_column_to_be_filtered_name] for index, row in source_file_data_frame.iterrows())
    unwanted_phrases_column = (row[unwanted_phrases_column_name] for index, row in unwanted_phrases_data_frame.iterrows())
    set_value = set(filtered_column) - set(unwanted_phrases_column)
    return pd.DataFrame(set_value).to_excel(filtered_file_path, header=False, index=False)

excel_editor(source_file_path, unwanted_phrases_file_path, filtered_file_path, source_file_column_to_be_filtered_name, unwanted_phrases_column_name)
