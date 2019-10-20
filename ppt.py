import comtypes.client
import sys
import os.path

powerpoint = comtypes.client.CreateObject("Powerpoint.Application")

def PPTtoPDF(inputFileName, outputFileName, formatType = 32):
    #powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    deck = powerpoint.Presentations.Open(inputFileName, WithWindow=False)
    deck.SaveAs(outputFileName, formatType) # formatType = 32 for ppt to pdf
    deck.Close()
    #powerpoint.Quit()



def main(inputFileName):
    print(inputFileName, " converting... ")
    outputFileName = inputFileName.split(".")[0] + ".pdf"

    absInput = os.path.abspath(inputFileName)
    absOutput = os.path.abspath(outputFileName)
    PPTtoPDF(absInput, absOutput)
    print("finish")

def main2():
    cwd = os.getcwd()
    file_list = os.listdir(cwd)

    for i in file_list:
        if i.endswith(".ppt") or i.endswith(".pptx"):
            main(i);

if __name__ == '__main__':
    if len(sys.argv) > 1:
        main(sys.argv[1])
    else:
        main2()

    powerpoint.Quit()

