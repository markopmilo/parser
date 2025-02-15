from sly import Lexer, Parser
import pandas as pd


class BibTexLexer(Lexer):
    tokens = {'ENTRY_KEY', 'KEY', 'VALUE', 'EQUALS'}
    ignore = ' \t'

    # Tokens
    ENTRY_KEY = r'@[a-zA-Z]+\{[^,]+'  # Capture @article{2024:
    KEY = r'[a-zA-Z_][a-zA-Z0-9_]*'  # Capture keys (author, orcid, journal, etc)
    VALUE = r'\{([^}]*)\}'  # The values for the keys (Душанић, 0000-0002..., књижевна историја, etc)
    EQUALS = r'='

    ignore_newline = r'\n+'
    ignore_comma = r',+'
    ignore_closing_brace = r'\}'

    def error(self, t):
        print("Illegal character '%s'" % t.value[0])
        self.index += 1


class BibTexParser(Parser):
    tokens = BibTexLexer.tokens

    # Set up empty dict of keys
    def __init__(self):
        self.keys = {}

    # capture @article line
    @_('ENTRY_KEY')
    def entry_key(self, p):
        self.entry_key = p.ENTRY_KEY.split('{')[1]

    # Analyze one line of bibliography
    @_('KEY EQUALS VALUE')
    def entry(self, p):
        value = p.VALUE[1:-1] if p.VALUE != '{}' else ''  # remove curly brace
        self.keys[p.KEY] = value

    # recursively analyze lines as shown above
    @_('entries entry')
    @_('entry')
    @_('entry_key')
    def entries(self, p):
        return self.keys

    # entry point
    @_('entries')
    def start(self, p):
        return self.keys


def load_excel_data(file_name, fields):
    df = pd.read_excel(file_name)  # reads data from file
    nameToId = dict(zip(df['author'], df['orcid']))
    idToName = dict(zip(df['orcid'], df['author'])) # assuming one-to-one relationship, I presume this would change, fine for now

    if not fields.get("orcid"):
        fields['orcid'] = nameToId.get(fields.get("author"))
    if not fields.get("author"):
        fields['author'] = idToName.get(fields.get("orcid"))


def format_bibtex(entry_key, fields):
    # Using the new fields dictionary, with the excel data, make a new string for the entry
    bibtex_string = f"@article{{{entry_key}, \n"
    for key, value in fields.items():
        bibtex_string += f"    {key} = {{{value}}},\n"
    bibtex_string += "}"
    return bibtex_string


def parse_bibtex(bibtex_entry, excel_file_name):
    # does the parsing work and returns the relevant data to the main method
    lexer = BibTexLexer()
    parser = BibTexParser()
    tokens = lexer.tokenize(bibtex_entry)
    fields = parser.parse(tokens)

    # add excel data if a field is missing
    if fields['orcid'] == "" or fields['author'] == "":
        load_excel_data(excel_file_name, fields)

    return parser.entry_key, fields


if __name__ == '__main__':
    bibtex_entry = """@article{2024:,
    author       = {Душанић, Дуња},
    showauthor   = {Разговор водила и с енглеског превела Дуња Душанић (12. март 2023.)},
    orcid        = {},
    title        = {После посткритике: разговор с Ритом Фелски},
    journal      = {Књижевна историја},
    year         = {2024},
    volume       = {56},
    number       = {183},
    pages        = {--},
    issn         = {0350-6428},
    howpublished = {print},
    language     = {serbian},
    note         = {7},
    doi          = {10.18485/kis.2024.56.183.7},
    doiurl       = {http://doi.fil.bg.ac.rs/volume.php?pt=journals&issue=kis-2024-56-183&i=7},
    url          = {http://doi.fil.bg.ac.rs/pdf/journals/kis/2024-183/kis-2024-56-183-7.pdf}
}
    """

    print("Previous BibTeX entry:")
    print(bibtex_entry)

    entry_info = parse_bibtex(bibtex_entry, "orcid.xlsx")

    print("\nUpdated BibTeX entry:")
    print(format_bibtex(entry_info[0], entry_info[1]))
