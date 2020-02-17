from pip._internal import main as pip
module_name = ["re","bs4","bleach","openpyxl","emoji","warnings","nltk","pathlib","sklearn"]
for i in module_name:
    try:
        __import__(i)
        print("Module Found : ",i)
    except:
        print("Do you wnat to install module : ["+i+"] (Y/N/E - Exit)")
        opt = input()
        if opt == "Y":
            pip(["install",i])
            print("Package install")
        elif opt == "N":
            print("Package skipped")
            pass
        elif opt == "E":
            exit()


