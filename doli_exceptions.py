class NumCompteNonRempli(Exception):
    """ Raise when column 1 'Num. Compta' is empty in template.xlsx sheet 'Compta' """
    pass

class DesignationNonRemplie(Exception):
    """ Raise when column 1 'Désignation' is empty in template.xlsx sheet 'Devis' """
    pass
class LienNonReconnu(Exception):
    """Raise when link enter by user is incorrect"""
    pass