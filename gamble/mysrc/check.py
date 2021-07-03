class Check:
    def average(self,seq):
        if len(self.info) > seq:
            return (self.sum(seq))/6
        else:
            return False

    def sum(self,seq):
        if len(self.info) > seq:
            return self.info[seq][0] + self.info[seq][1] + self.info[seq][2] + self.info[seq][3] + self.info[seq][4] + self.info[seq][5]
        else:
            return False

    def __del__(self):
        pass

    def __new__(self,ExternalInfo):
        return super().__new__(self)

    def __init__(self,ExternalInfo):
        self.info = ExternalInfo
        super().__init__()
    


if __name__ == '__main__':
    print('This is check.py, Do execute main.py')