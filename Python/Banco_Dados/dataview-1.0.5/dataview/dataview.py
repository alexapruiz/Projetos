from collections import Sequence


class DataView(Sequence):
    def __init__(self, data, *args):
        self.data = data

        slice_ = slice(*args) if args else slice(None, None, None)

        self.start, self.stop, self.step = slice_.start, slice_.stop, slice_.step

    @property
    def step(self):
        return self.__step

    @step.setter
    def step(self, value):
        if value == 0:
            raise ValueError("DataView step cannot be zero")

        self.__step = value

    @property
    def indices(self):
        return slice(self.start, self.stop, self.step).indices(len(self.data))

    def __getitem__(self, item):
        if type(item) == slice:
            if item.step == 0:
                raise ValueError("slice step cannot be zero")

            return DataView(self, item.start, item.stop, item.step)

        elif type(item) == int:
            if item in range(len(self)):
                return self.data[self.indices[0] + item * self.indices[2]]
            elif item in range(-len(self), 0):
                return self.data[self.indices[0] + (item + len(self)) * self.indices[2]]
            else:
                raise IndexError("DataView index out of range")

        else:
            raise TypeError("DataView indices should be integer")

    def __len__(self):
        return len(range(*self.indices))

    def __repr__(self):
        return str(list(self))
