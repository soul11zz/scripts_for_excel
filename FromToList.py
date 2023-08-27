

class FromToList:
    _ftl = {}

    def _if_exist(self ,dataset ,key):
        return key in dataset

    def write(self, key: str, collection: str):
        key = int(key)
        collection = int(collection)
        if self._if_exist(self._ftl, key):
            if not self._if_exist(self._ftl[key], collection):
                self._ftl[key].add(collection)
                return True
        else:
            self._ftl[key] = {collection}
            return True
        return False

    def check(self, key: str, collection: str):
        key = int(key)
        collection = int(collection)
        if self._if_exist(self._ftl, key):
            if not self._if_exist(self._ftl[key], collection):
                return True
        else:
            return True
        return False

    def print(self):
        print(self._ftl)

    def return_list(self):
        res = []
        for key, values in self._ftl.items():
            for value in values:
                res.append([key, value])

        return res


