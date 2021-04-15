class Anuncio:

    def __init__(self, tipo, recurso, texto, fecha, autorimagen, autortexto):
        self.tipo = tipo
        self.recurso = recurso
        self.texto = texto
        self.fecha = fecha
        self.autorimagen = autorimagen
        self.autortexto = autortexto
    
    def getTipo(self):
        return self.tipo

    def getMedia(self):
        return self.recurso


    def getFecha(self):
        return self.fecha

    def getDescription(self):
        return self.texto

    def getAuthMedia(self):
        return self.autorimagen
    
    def getAuthTexto(self):
        return self.autortexto