import Datasheet_Download
import  preprocessing
import RSI

toDownload  = input("Do you want to downlaod the historical data (y/n) :")
toContinue =  "y"
if (toDownload.lower() == "y"):
    Datasheet_Download.Download_Dataheet()
    toContinue =  input("Do you want to continue(y/n) : ")
if (toContinue  == "y"):
    preprocessing.preprocess()
    RSI.Execute()