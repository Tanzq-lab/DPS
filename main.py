from Logic.Function import *

if __name__ == '__main__':
    infoFactory = InfoFactory()
    FileTypeDict = init()

    while True:
        opt = menu()
        if opt == 0:
            break
        elif opt == 1:
            show(FileTypeDict)
        elif opt == 2:
            query(FileTypeDict)
        elif opt == 3:
            dataProcessing(infoFactory, FileTypeDict)

    print("感谢您的使用!(●'◡'●)")
