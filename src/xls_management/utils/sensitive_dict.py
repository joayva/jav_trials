class SensitiveDict(dict):
    def __setitem__(self, key, value):
        """
            key colisions are resolved by keeping the first value
        """
        if key in self.keys():
            raise KeyError(message=f"{key} already exists")
        super().__setitem__(key, value)