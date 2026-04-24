class DictFirst(dict):
    def __setitem__(self, key, value):
        """
            key colisions are resolved by keeping the first value
        """
        if key not in self.keys():
            super().__setitem__(key, value)