class FillDict(dict):
    def __setitem__(self, key, value):
        assert isinstance(value,dict)
        if key not in self.keys():
            super[key]=value
        else:
            for k, v in value:
                super[key][k]=v
    
    def __getitem__(self, key):
        if key not in self.keys():
            super().__setitem__(key,{})
        return super().__getitem__(key)
