import csv
from pathlib import Path
from tkinter.filedialog import askopenfilename, asksaveasfilename
from tkinter import messagebox

from codesys_symbols_parser import CodesysSymbolParser
from xls_write import write_xls


def ask_for_overwrite(filepath):
    """If file exists, prompt for overwrite or to select another filepath"""
    if filepath.exists():
        overwrite = messagebox.askquestion(
            "Overwrite existing target file",
            f"The selected target file {filepath} already exists on your filesystem.\n\n"
            f"Press 'YES' if you wish to overwrite it or 'NO' to select another filename",
            type=messagebox.YESNO)

        if overwrite == messagebox.NO:
            return asksaveasfilename(filetypes=[('CSV files', '.csv'), ('All files', '.*')])
    return filepath


def main():
    symbols_filepath = askopenfilename(title='Please choose a CoDeSys application symbols file to open',
                                       filetypes=[('XML files', '.xml'), ('All files', '.*')])
    if not symbols_filepath:
        messagebox.showerror("No input file selected",
                             "No file selected.")
        return -1

    parser = CodesysSymbolParser(symbols_filepath)
    parser.parse()
    symbols = parser.get_symbols()
    csv_out_filepath = ask_for_overwrite(Path(symbols_filepath).with_suffix('.csv'))
    if not csv_out_filepath:
        messagebox.showerror("No output file selected",
                             "No file selected to save the results.")
        return -1

    fieldnames = ['name', 'comment']
    with open(csv_out_filepath, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writerows(symbols)

    xlsx_out_filepath = ask_for_overwrite(Path(symbols_filepath).with_suffix('.xlsx'))
    write_xls(xlsx_out_filepath, symbols)

    messagebox.showinfo("Symbols saved",
                        f'{len(symbols)} symbols found.\n'
                        f'File saved to :\n'
                        f'{csv_out_filepath}\n'
                        f'{xlsx_out_filepath}')


if __name__ == '__main__':
    main()
