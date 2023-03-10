class AnaliticoDTO:
    
     def __init__(self, desp, nomedesp, data, credor, discrimina, codigo, valor, PORCENTO, TOTAL):
        self.desp  = desp
        self.nomedesp = nomedesp
        self.data = data
        self.credor = credor
        self.discrimina = discrimina
        self.codigo = codigo
        self.valor = valor
        self.PORCENTO = PORCENTO
        self.TOTAL = TOTAL
        self.TotalDespesas = None
    
pass