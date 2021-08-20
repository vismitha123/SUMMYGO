
import win32com.client


def PPTtoPDF():
    print("reached-2")
    inputFileName= r"C:\Users\lenovo\Desktop\FYP\backend\Output Presentations\OUTPUT.pptx"
    outputFileName= r"C:\Users\lenovo\Desktop\FYP\backend\Output Presentations\Output.pdf"
    formatType = 32
    print("reached-3")
    #powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint = win32com.client.DispatchEx("Powerpoint.Application")
    powerpoint.Visible = 1
    print("reached-4")
    if outputFileName[-3:] != 'pdf':
        outputFileName = outputFileName + ".pdf"
    deck = powerpoint.Presentations.Open(inputFileName)
    deck.SaveAs(outputFileName, formatType) # formatType = 32 for ppt to pdf
    print("reached-5")
    deck.Close()
    powerpoint.Quit()


PPTtoPDF()
