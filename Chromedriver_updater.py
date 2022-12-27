import os, zipfile
from selenium import webdriver
from win32com.client import Dispatch

def config():
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('log-level=3')
    chrome_options.add_argument('--headless')
    return chrome_options

def get_version_via_com(filename):
    parser = Dispatch("Scripting.FileSystemObject")
    try:
        version = parser.GetFileVersion(filename)
    except Exception:
        return None
    return version

def get_version():
    paths = [r"C:\Program Files\Google\Chrome\Application\chrome.exe",
             r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"]
    version = list(filter(None, [get_version_via_com(p) for p in paths]))[0]
    return version

def format_version(version):
    aux = version[::-1]
    index = -aux.find('.')
    return int(version[index:]), version[:index]

def install_driver():
    version = get_version()
    link = "https://chromedriver.storage.googleapis.com/"
    i, version = format_version(version)

    resp = 1
    while resp == 1 or i<0:
        path = 'chromedriver'
        os.system("curl "+link+version+str(i)+"/chromedriver_win32.zip -s --output chromedriver.zip")

        try:
            zip_ref = zipfile.ZipFile(path+".zip", 'r')
            zip_ref.extractall(path)
            resp = 0
        except zipfile.BadZipFile as error:
            pass
        i-=1

    if resp == 0:
        print(f"Driver ({version+str(i+1)}) instaled sucesseful\n")
    else:
        print(f"Driver can\'t be found\n")

def move():
    os.rename("chromedriver/chromedriver.exe", 'chromedriver.exe')
    os.rmdir("chromedriver")

def main():
    os.chdir(os.path.dirname(__file__)+"/drivers")
    try:
        web = webdriver.Chrome("chromedriver.exe", options=config())
        print("Driver alredy instaled")
        web.close()
    except:
        install_driver()
        move()

if __name__ == '__main__':
    main()