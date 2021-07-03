class GetHistory:
    def __del__(self):
        pass

    def __new__(self,ExternalInfo):
        return super().__new__(self)

    def __init__(self,ExternalInfo):
        self.info = ExternalInfo
        super().__init__()
    
if __name__ == '__main__':
    print('This is gethistory.py, Do execute main.py')
