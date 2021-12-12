"""
In questo file ho sia del codice che viene eseguito:
print("Hai lanciato lo script da questo stesso file")

ma ho anche dei parametri e delle funzioni che possono servirvi: stocazzo e foo.
Se io voglio accedere a questi parametri ma non voglio eseguire le istruzioni come devo fare?
Metto il codice sotto la flag:
if __name__ == '__main__':

le varianili precedute e finite con doppio underscore sono "variabili speciali" che servono ad essere lette ma non saranno mai usate.
La variabile __name__ assume valore "__mail__" se lo script viene eseguito come principale.

"""
stocazzo = 1
def foo()->None:
    print("Sei entrato nella funzione foo!")


# Se io voglio importa
if __name__ == '__main__':
    print("Hai lanciato lo script da questo stesso file")