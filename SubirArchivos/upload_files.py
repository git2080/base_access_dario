import pandas as pd
import urllib.request as ul
import zipfile as zf
import numpy as np
import shutil
from GoogleDrive import authentication, update_file
import httplib2

def main():
    url = np.array([
        ["OEC.zip", "http://10.232.73.202/pe/sred/planillas/OEC.zip"],
        ["UDPC.zip", "http://10.232.73.202/pe/sred/planillas/UDPC.zip"],
        ["ATP.zip", "http://10.232.73.202/pe/atp/planillas/ATP.zip"]])
    for i in range (len(url)):
        str_url1 = str(url[i][0])
        str_url2 = str(url[i][1])
        ul.urlretrieve(str_url2, str_url1)
        shutil.unpack_archive(str_url1)
    update_file(authentication(), 'ATP.mdb', 'application/x-msaccess', "1nwNLakuC3Acjke0g2Upa1MBcg84YpeIO")
    update_file(authentication(), 'OEC.mdb', 'application/x-msaccess', "1R1kcA64LPipJ7a37shMcEbhkwYkjA1TJ")
    update_file(authentication(), 'MATOEC.mdb', 'application/x-msaccess', "12t6V3Vr_E4dHjAlb30EMrczkVZM3iTl4")
main()