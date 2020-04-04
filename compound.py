# Compound class storing ID/name of a compound, its type, and its IC50 val

class Compound:

    # Compound constructor
    def __init__(self, compoundID, compoundType, compoundIC):
        self.id = compoundID
        self.type = compoundType
        self.ic = compoundIC
    
    # Return compound ID/name
    def getID(self):
        return self.id

    # Return type of compound
    def getType(self):
        return self.type

    # Return IC50 val of compound
    def getIC(self):
        return self.ic
